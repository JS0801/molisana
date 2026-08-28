/*eslint-disable*/
/**
 * Inbound Email Plugin:
 * - Reads CSV from email attachment
 * - If attachment type is CSV / PLAINTEXT => use as-is (NO base64 decode)
 * - If attachment type is MISC (or unknown) => try base64 decode
 * - Parses CSV safely (supports commas inside quotes)
 * - Creates customrecord_mi_planned_po records
 */

function process(email) {
  try {
    var subject = email.getSubject();
    var sender = email.getFrom();
    nlapiLogExecution("DEBUG", "Email Received", "Subject: " + subject + ", Sender: " + sender);

    var attachments = email.getAttachments();
    nlapiLogExecution("AUDIT", "Email Attachments", "Count: " + (attachments ? attachments.length : 0));

    if (!attachments || attachments.length === 0) {
      nlapiLogExecution("DEBUG", "No Attachments", "No attachments found in the email.");
      return;
    }

    // -----------------------------
    // 1) Pick CSV attachment
    // -----------------------------
    var csvContent = null;
    var pickedName = null;
    var pickedType = null;

    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];

      var fileName = (att.getName && att.getName()) ? att.getName() : "";
      var fileType = (att.getType && att.getType()) ? String(att.getType()) : "";

      nlapiLogExecution("DEBUG", "Attachment Found", "Name: " + fileName + ", Type: " + fileType);

      // Prefer filename because inbound attachment type may be MISC
      var isCsvByName = fileName && fileName.toLowerCase().indexOf(".csv") !== -1;

      if (isCsvByName || i === 0) {
        pickedName = fileName || ("attachment_" + (i + 1) + ".csv");
        pickedType = fileType || "";

        if (att.getValue) csvContent = att.getValue();
        else if (att.getContents) csvContent = att.getContents();
        else csvContent = null;

        break;
      }
    }

    if (!csvContent) {
      nlapiLogExecution("ERROR", "No CSV Found", "Could not read attachment content.");
      return;
    }

    // -----------------------------
    // 2) Decode ONLY when needed
    //    If fileType is CSV (or PLAINTEXT), DO NOT base64 decode.
    //    If fileType is MISC/unknown, try base64 decode.
    // -----------------------------
    var rawPreview = safePreview(csvContent, 120);
    nlapiLogExecution("DEBUG", "Attachment Preview (raw)", "Name: " + pickedName + " Type: " + pickedType + " Preview: " + rawPreview);

    var typeUpper = String(pickedType || "").toUpperCase();

    // If NetSuite says CSV => treat as plain text, no decode
    // (Same for PLAINTEXT)
    var isPlainTextType = (typeUpper === "CSV" || typeUpper === "PLAINTEXT" || typeUpper === "TEXT");

    if (!isPlainTextType) {
      // Only attempt decode for MISC/unknown types
      if (looksBase64(csvContent)) {
        nlapiLogExecution("AUDIT", "Attachment Decode", "Type not CSV/TEXT. Looks like base64. Decoding. File: " + pickedName);
        csvContent = base64Decode(csvContent);
        nlapiLogExecution("DEBUG", "Attachment Preview (decoded)", safePreview(csvContent, 120));
      }
    } else {
      nlapiLogExecution("DEBUG", "Attachment Decode", "Skipped decode because type is " + typeUpper);
    }

    // Guard: If content is still not readable CSV text, stop early
    if (!looksLikeCsvText(csvContent)) {
      nlapiLogExecution(
        "ERROR",
        "Invalid CSV Content",
        "Attachment content is not readable CSV text. File: " + pickedName +
          " Type: " + pickedType +
          " Preview: " + safePreview(csvContent, 120)
      );
      return;
    }

    // -----------------------------
    // 3) Parse CSV (quote-safe)
    // -----------------------------
    var rows = parseCSV(csvContent);
    if (!rows || rows.length === 0) {
      nlapiLogExecution("ERROR", "Empty CSV", "CSV content appears to be empty or unparseable.");
      return;
    }

    nlapiLogExecution("DEBUG", "CSV Rows", "Total Rows: " + rows.length);

    // -----------------------------
    // 4) Build header map
    // -----------------------------
    var headers = rows[0];
    var headerMap = {};
    nlapiLogExecution("DEBUG", "CSV Headers", JSON.stringify(headers));

    for (var h = 0; h < headers.length; h++) {
      var headerName = (headers[h] !== null && headers[h] !== undefined) ? String(headers[h]).trim().toLowerCase() : "";

      if (headerName === "preferred vendor id") headerMap.vendor = h;
      if (headerName === "item id") headerMap.item = h;
      if (headerName === "description") headerMap.description = h;
      if (headerName === "comments") headerMap.comments = h;
      if (headerName === "order calculator") headerMap.qty = h;
      if (headerName === "month of stock") headerMap.monthstock = h;
      if (headerName === "4 month average") headerMap.avg4qty = h;
      if (headerName === "in transit") headerMap.transit = h;
      if (headerName === "on order (to ship)") headerMap.onorder = h;
      if (headerName === "total available") headerMap.avail = h;
      if (headerName === "zone") headerMap.zone = h;
    }

    nlapiLogExecution("DEBUG", "Header Map", JSON.stringify(headerMap));

    if (headerMap.vendor === undefined || headerMap.item === undefined) {
      nlapiLogExecution("ERROR", "Missing Required Headers", "Required: Preferred Vendor ID, Item ID");
      return;
    }

    // -----------------------------
    // 5) Create records
    // -----------------------------
    var created = 0;
    var failed = 0;

    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      if (!row || row.length < 2) continue;

      var vendorVal = getCell(row, headerMap.vendor);
      var itemVal = getCell(row, headerMap.item);
      if (isEmpty(vendorVal) || isEmpty(itemVal)) continue;

      try {
        var newRecord = nlapiCreateRecord("customrecord_mi_planned_po");

        newRecord.setFieldValue("custrecord_mi_preffered_vendor", String(vendorVal).trim());
        newRecord.setFieldValue("custrecord_mi_item", String(itemVal).trim());

        var descVal = getCell(row, headerMap.description);
        if (!isEmpty(descVal)) newRecord.setFieldValue("custrecord_item_name", String(descVal).trim());

        var qtyVal = getCell(row, headerMap.qty);
        if (!isEmpty(qtyVal)) newRecord.setFieldValue("custrecord_mi_order_qty", toNumberString(qtyVal));

        var monthVal = getCell(row, headerMap.monthstock);
        if (!isEmpty(monthVal)) newRecord.setFieldValue("custrecord_month_of_stocks", toNumberString(monthVal));

        var avg4Val = getCell(row, headerMap.avg4qty);
        if (!isEmpty(avg4Val)) newRecord.setFieldValue("custrecord_mi_min_month_qty", toNumberString(avg4Val));

        var transitVal = getCell(row, headerMap.transit);
        if (!isEmpty(transitVal)) newRecord.setFieldValue("custrecord_mi_qty_in_transit", toNumberString(transitVal));

        var onOrderVal = getCell(row, headerMap.onorder);
        if (!isEmpty(onOrderVal)) newRecord.setFieldValue("custrecord_mi_qty_on_order_not_recv", toNumberString(onOrderVal));

        var availVal = getCell(row, headerMap.avail);
        if (!isEmpty(availVal)) newRecord.setFieldValue("custrecord_mi_qty_available", toNumberString(availVal));

        var commentVal = getCell(row, headerMap.comments);
        if (!isEmpty(commentVal)) newRecord.setFieldValue("custrecord_mi_purchase_memo", String(commentVal).trim());

        newRecord.setFieldValue("custrecord_mi_approval_status", "1");

        var recId = nlapiSubmitRecord(newRecord, true, true);
        created++;
        nlapiLogExecution("AUDIT", "Record Created", "Row: " + r + " ID: " + recId);

      } catch (e) {
        failed++;
        nlapiLogExecution("ERROR", "Error Creating Record (Row " + r + ")", e.toString());
      }
    }

    nlapiLogExecution("AUDIT", "Process Complete", "Created: " + created + ", Failed: " + failed);

  } catch (error) {
    nlapiLogExecution("ERROR", "Critical Error", error.toString());
  }
}

/* =========================
 * Helpers
 * ========================= */

function isEmpty(v) {
  return v === null || v === undefined || String(v).trim() === "";
}

function getCell(row, idx) {
  if (idx === undefined || idx === null) return "";
  if (!row || idx >= row.length) return "";
  return row[idx];
}

function toNumberString(v) {
  var s = String(v);
  s = s.replace(/,/g, "").replace(/\s+/g, "");
  if (s === "") return "";
  return s;
}

function safePreview(s, n) {
  try {
    s = String(s);
    if (s.length <= n) return s;
    return s.substr(0, n) + "...";
  } catch (e) {
    return "";
  }
}

function looksLikeCsvText(content) {
  if (!content) return false;
  var s = String(content);

  if (s.length >= 2 && s.charAt(0) === "P" && s.charAt(1) === "K") return false;

  var hasComma = s.indexOf(",") !== -1;
  var hasNewline = s.indexOf("\n") !== -1 || s.indexOf("\r") !== -1;

  var printable = 0;
  var limit = Math.min(200, s.length);
  for (var i = 0; i < limit; i++) {
    var code = s.charCodeAt(i);
    if (code === 9 || code === 10 || code === 13 || (code >= 32 && code <= 126)) printable++;
  }
  var printableRatio = limit ? (printable / limit) : 0;

  return hasComma && hasNewline && printableRatio > 0.85;
}

function looksBase64(s) {
  if (!s) return false;
  s = String(s).replace(/\s+/g, "");
  if (s.length < 50) return false;
  return /^[A-Za-z0-9+/=]+$/.test(s);
}

function base64Decode(b64) {
  var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var str = String(b64).replace(/[\r\n\s]/g, "");
  var output = "";
  var buffer = 0, bits = 0;

  for (var i = 0; i < str.length; i++) {
    var c = str.charAt(i);
    if (c === "=") break;
    var val = chars.indexOf(c);
    if (val < 0) continue;

    buffer = (buffer << 6) | val;
    bits += 6;

    if (bits >= 8) {
      bits -= 8;
      output += String.fromCharCode((buffer >> bits) & 0xff);
    }
  }
  return output;
}

/* =========================
 * CSV Parser (quote-safe)
 * ========================= */

function parseCSV(content) {
  var lines = String(content).split(/\r\n|\n/);
  var result = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!line || String(line).trim() === "") continue;
    result.push(parseCSVLine(line));
  }
  return result;
}

function parseCSVLine(line) {
  var out = [];
  var cur = "";
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var ch = line.charAt(i);

    if (ch === '"') {
      if (inQuotes && i + 1 < line.length && line.charAt(i + 1) === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === "," && !inQuotes) {
      out.push(cur);
      cur = "";
    } else {
      cur += ch;
    }
  }
  out.push(cur);

  for (var j = 0; j < out.length; j++) {
    out[j] = (out[j] === null || out[j] === undefined) ? "" : String(out[j]).trim();
  }
  return out;
}
