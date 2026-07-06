/*eslint-disable*/
/**
 * Inbound Email Plugin (SuiteScript 1.0)
 * - Reads email attachments (Proforma Invoice / any file type)
 * - Finds PO ID from subject token: [POID:12345]
 * - If attachment exists:
 *    - Saves file to File Cabinet (adds 4-char random string in name)
 *    - Sets file description = "Proforma Invoice"
 *    - Attaches file to Purchase Order
 *    - Sends confirmation email
 * - If no attachment:
 *    - Sends auto-reply email with "Reply with Invoice" mailto button (prefilled subject/body)
 *
 * IMPORTANT FIX:
 * - Avoids "Script Execution Instruction Count Exceeded" by using native base64 decode:
 *   nlapiDecrypt('base64', ...)
 */

function process(email) {
  try {
    var subject = email.getSubject() || '';
    var sender = email.getFrom() || '';
    nlapiLogExecution("AUDIT", "Inbound Email Received", "Subject: " + subject + " | From: " + sender);

    var poId = extractPoId(subject);
    if (!poId) {
      nlapiLogExecution("ERROR", "Missing POID", "Subject must contain [POID:12345]. Subject: " + subject);
      return;
    }

    var attachments = email.getAttachments();
    nlapiLogExecution("AUDIT", "Email Attachments", "Count: " + (attachments ? attachments.length : 0));

    // -----------------------------
    // 1) No attachment => auto reply
    // -----------------------------
    if (!attachments || attachments.length === 0) {
      nlapiLogExecution("DEBUG", "No Attachments", "Sending auto-reply. POID: " + poId);
      sendMissingAttachmentReply(sender, subject, poId);
      return;
    }

    // -----------------------------
    // 2) Save & attach each file
    // -----------------------------
    var folderId = 442364; // your File Cabinet folder internal id
    var attachedCount = 0;

    // Safety limits (avoid huge workload)
    var MAX_ATTACHMENTS = 5;
    var MAX_CONTENT_LEN = 12 * 1024 * 1024; // ~12MB chars (adjust if needed)

    for (var i = 0; i < attachments.length; i++) {
      if (i >= MAX_ATTACHMENTS) {
        nlapiLogExecution("AUDIT", "Attachment Limit", "Skipping extra attachments. Max: " + MAX_ATTACHMENTS);
        break;
      }

      var att = attachments[i];

      var originalName = (att.getName && att.getName()) ? att.getName() : ("attachment_" + (i + 1));
      var fileType = (att.getType && att.getType()) ? String(att.getType()) : "MISC";

      var rand = random4();
      var fileName = poId + "_" + rand + "_" + originalName;

      nlapiLogExecution("DEBUG", "Attachment Found", "Name: " + fileName + " | Type: " + fileType);

      // Get contents
      var contents = null;
      if (att.getValue) contents = att.getValue();
      else if (att.getContents) contents = att.getContents();

      if (!contents) {
        nlapiLogExecution("ERROR", "Attachment Read Failed", "Could not read contents. Name: " + fileName);
        continue;
      }

      var rawLen = String(contents).length;
      nlapiLogExecution("DEBUG", "Attachment Content Length (raw)", "Name: " + fileName + " | Len: " + rawLen);

      if (rawLen > MAX_CONTENT_LEN) {
        nlapiLogExecution("ERROR", "Attachment Too Large", "Skipping file. Name: " + fileName + " | Len: " + rawLen);
        continue;
      }

      // Decide if decode is needed
      var typeUpper = String(fileType || "").toUpperCase();
      var isPlainTextType = (typeUpper === "CSV" || typeUpper === "PLAINTEXT" || typeUpper === "TEXT");

      // var lb64 = looksBase64(contents);

      // nlapiLogExecution("DEBUG", "Decode Decision", JSON.stringify({
      //   name: fileName,
      //   type: fileType,
      //   isPlainTextType: isPlainTextType,
      //   looksBase64: lb64,
      //   len: rawLen
      // }));

      // // Decode ONLY if not plain text AND looks base64
      // if (!isPlainTextType && lb64) {
      //   nlapiLogExecution("AUDIT", "Decoding Attachment", "Using native base64 decode. Name: " + fileName);
      //   var decoded = decodeBase64Native(contents);

      //   if (decoded) {
      //     contents = decoded;
      //     nlapiLogExecution("DEBUG", "Decoded OK", "Name: " + fileName + " | DecodedLen: " + String(contents).length);
      //   } else {
      //     nlapiLogExecution("ERROR", "Decode Failed", "Skipping file because decode returned null. Name: " + fileName);
      //     continue;
      //   }
      // }

      // Create file in File Cabinet
      var nsFileType = mapInboundTypeToNsFileType(fileType, fileName);
      nlapiLogExecution("DEBUG", "File Type Mapping", "Inbound: " + fileType + " => NS: " + nsFileType + " | Name: " + fileName);

      var f = nlapiCreateFile(fileName, nsFileType, contents);
      f.setFolder(folderId);
      f.setDescription('Proforma Invoice');

      var fileId = nlapiSubmitFile(f);
      nlapiLogExecution("AUDIT", "File Saved", "FileId: " + fileId + " | Name: " + fileName);

      // Attach to PO
      nlapiAttachRecord("file", fileId, "purchaseorder", poId);
      attachedCount++;

      nlapiLogExecution("AUDIT", "File Attached to PO", "POID: " + poId + " | FileId: " + fileId);
    }

    // If attachments existed but none saved/attached for some reason -> reply
    if (attachedCount === 0) {
      nlapiLogExecution("ERROR", "No Valid Attachments", "Attachments existed but none were processed. Sending reply. POID: " + poId);
      sendMissingAttachmentReply(sender, subject, poId);
      return;
    }

    // Confirmation reply back to sender
    sendReceivedConfirmation(sender, subject, poId, attachedCount);

    nlapiLogExecution("AUDIT", "Process Complete", "POID: " + poId + " | Attached Files: " + attachedCount);

  } catch (e) {
    nlapiLogExecution("ERROR", "Critical Error", e.toString());
  }
}

/* =========================
 * Helpers
 * ========================= */

function extractPoId(subject) {
  var m = String(subject || '').match(/\[POID:(\d+)\]/i);
  return m ? m[1] : null;
}

function random4() {
  var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  var out = "";
  for (var i = 0; i < 4; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

function escHtml(s) {
  s = (s == null ? '' : String(s));
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function urlEncode(s) {
  return encodeURIComponent(String(s || ''));
}

function buildReplyMailtoHref(replyToEmail, subjectWithPoid) {
  var body =
    "Hi,%0D%0A%0D%0A" +
    "Please find attached Proforma Invoice (mandatory) as proof of acknowledgement for this PO.%0D%0A%0D%0A" +
    "Thanks%0D%0A";

  return "mailto:" + replyToEmail +
         "?subject=" + urlEncode(subjectWithPoid) +
         "&body=" + body;
}

function sendMissingAttachmentReply(toEmail, origSubject, poId) {
  if (!toEmail) return;

  // Reply back to your NetSuite inbound email address (so it comes back into NetSuite)
  var replyToEmail = "emails.4975346.2753.98b7678d92@4975346.email.netsuite.com";

  var subjBase = String(origSubject || '');
  var subj = (subjBase.toLowerCase().indexOf('re:') === 0) ? subjBase : ("Re: " + subjBase);

  // Ensure POID token exists silently (no visible "Important" text)
  if (!/\[POID:\d+\]/i.test(subj)) subj += " [POID:" + poId + "]";

  var mailtoHref = buildReplyMailtoHref(replyToEmail, subj);

  var htmlBody = ''
    + '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.45;">'
    + '  <p>Hi,</p>'
    + '  <p>We received your PO acknowledgement email, but <b>no attachment</b> was included.</p>'
    + '  <p style="margin:10px 0 16px 0;">'
    + '    <span style="font-size:18px;">📎</span> Please send the <b>Proforma Invoice</b> (mandatory) as proof of acknowledgement.'
    + '  </p>'
    + '  <p style="margin:18px 0;">'
    + '    <a href="' + escHtml(mailtoHref) + '" '
    + '       style="display:inline-block;padding:12px 18px;border-radius:10px;'
    + '              background:#2563eb;color:#fff;text-decoration:none;font-weight:bold;">'
    + '      Reply with Invoice'
    + '    </a>'
    + '  </p>'
    + '  <p style="color:#666;font-size:12px;margin-top:10px;">'
    + '    If the button does not work, reply to this email and attach the Proforma Invoice.'
    + '  </p>'
    + '  <p>Thanks</p>'
    + '</div>';

  // Send HTML email (works in most accounts)
  nlapiSendEmail(-5, toEmail, subj, htmlBody, null, null, null, null, true);
}

function sendReceivedConfirmation(toEmail, origSubject, poId, count) {
  if (!toEmail) return;

  var body =
    "Hi,\n\n" +
    "Thank you. We received the Proforma Invoice attachment(s) and linked them to the Purchase Order.\n\n" +
    "PO ID: " + poId + "\n" +
    "Attachments received: " + count + "\n\n" +
    "Thanks";

  var subj = (String(origSubject || '').toLowerCase().indexOf('re:') === 0) ? origSubject : ("Re: " + origSubject);

  nlapiSendEmail(-5, toEmail, subj, body);
}

function mapInboundTypeToNsFileType(inboundType, fileName) {
  var t = String(inboundType || '').toUpperCase();
  var name = String(fileName || '').toLowerCase();

  // Prefer extension match
  if (name.indexOf('.pdf') !== -1) return 'PDF';
  if (name.indexOf('.png') !== -1) return 'PNGIMAGE';
  if (name.indexOf('.jpg') !== -1 || name.indexOf('.jpeg') !== -1) return 'JPGIMAGE';
  if (name.indexOf('.doc') !== -1 || name.indexOf('.docx') !== -1) return 'WORD';
  if (name.indexOf('.xls') !== -1 || name.indexOf('.xlsx') !== -1) return 'EXCEL';
  if (name.indexOf('.csv') !== -1) return 'CSV';
  if (name.indexOf('.txt') !== -1) return 'PLAINTEXT';

  // Fallback on inbound type if it matches known values
  if (t === 'PDF') return 'PDF';
  if (t === 'WORD') return 'WORD';
  if (t === 'EXCEL') return 'EXCEL';
  if (t === 'CSV') return 'CSV';
  if (t === 'PLAINTEXT' || t === 'TEXT') return 'PLAINTEXT';

  // Generic fallback for any unknown types
  return 'MISC';
}

function looksBase64(s) {
  if (!s) return false;
  s = String(s).replace(/\s+/g, "");
  if (s.length < 50) return false;
  return /^[A-Za-z0-9+/=]+$/.test(s);
}

// Native decode to avoid "Instruction Count Exceeded"
function decodeBase64Native(b64) {
  try {
    return nlapiDecrypt('base64', String(b64 || '').replace(/[\r\n\s]/g, ''));
  } catch (e) {
    nlapiLogExecution('ERROR', 'Base64 decode failed', e.toString());
    return null;
  }
}