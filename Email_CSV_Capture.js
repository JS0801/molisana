/*eslint-disable*/
/**
 * Inbound Email Plugin
 * Email CSV -> customrecord_ds_price_level_change_reque
 */

var TARGET_RECORD = 'customrecord_ds_price_level_change_reque';

function process(email) {
  try {
    var subject = email.getSubject();
    var sender = email.getFrom();

    nlapiLogExecution('DEBUG', 'Email Received', 'Subject: ' + subject + ', Sender: ' + sender);

    var attachments = email.getAttachments();

    if (!attachments || attachments.length === 0) {
      nlapiLogExecution('ERROR', 'No Attachments', 'No attachment found.');
      return;
    }

    var csvContent = null;
    var pickedName = '';

    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];
      var fileName = att.getName ? att.getName() : '';

      if (fileName && fileName.toLowerCase().indexOf('.csv') !== -1) {
        pickedName = fileName;

        if (att.getValue) {
          csvContent = att.getValue();
        } else if (att.getContents) {
          csvContent = att.getContents();
        }

        break;
      }
    }

    if (!csvContent) {
      nlapiLogExecution('ERROR', 'CSV Missing', 'Could not read CSV content.');
      return;
    }

    if (looksBase64(csvContent)) {
      csvContent = base64Decode(csvContent);
    }

    nlapiLogExecution('DEBUG', 'CSV File Picked', pickedName);
    nlapiLogExecution('DEBUG', 'CSV Preview', String(csvContent).substr(0, 1000));

    var rows = parseCSV(csvContent);

    if (!rows || rows.length <= 1) {
      nlapiLogExecution('ERROR', 'CSV Empty', 'No data rows found.');
      return;
    }

    var headerMap = buildHeaderMap(rows[0]);

    var created = 0;
    var failed = 0;

    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];

      if (!row || row.length === 0 || isBlankRow(row)) {
        continue;
      }

      try {
        nlapiLogExecution('DEBUG', 'Processing Row ' + r, JSON.stringify(row));

        var rec = nlapiCreateRecord(TARGET_RECORD);

        // custrecord_ds_start_date <- Accepted Start Date
        setDateValue(rec, 'custrecord_ds_start_date', getByHeader(row, headerMap, 'custrecord_ds_start_date'));

        setDateValue(rec, 'custrecord_ds_end_date', getByHeader(row, headerMap, 'custrecord_ds_end_date'));

        // custrecord_ds_price_currency <- always Canadian Dollar
        setSelectValue(rec, 'custrecord_ds_price_currency', 'Canadian Dollar', 'currency');

        setTextOrValue(rec, 'custrecord_ds_old_price', cleanNumber(getByHeader(row, headerMap, 'custrecord_ds_old_price')));

        // custrecord_ds_new_price <- Accepted Price
        setTextOrValue(rec, 'custrecord_ds_new_price', cleanNumber(getByHeader(row, headerMap, 'custrecord_ds_new_price')));

        setSelectValue(rec, 'custrecord_ds_price_level', getByHeader(row, headerMap, 'custrecord_ds_price_level'), 'pricelevel');

        setSelectValue(rec, 'custrecord_ds_price_level_item', getByHeader(row, headerMap, 'custrecord_ds_price_level_item'), 'item');

        setTextOrValue(rec, 'custrecord160', getByHeader(row, headerMap, 'custrecord160'));

        setCheckboxValue(rec, 'custrecordis_lot', getByHeader(row, headerMap, 'custrecordis_lot'));

        // This will populate by internal ID if CSV value is numeric
        setSelectValue(rec, 'custrecord_new_price_level', getByHeader(row, headerMap, 'custrecord_new_price_level'), 'pricelevel');

        setCustomerValue(rec, 'custrecord_mi_customers', getByHeader(row, headerMap, 'custrecord_mi_customers'));

        setCheckboxValue(rec, 'custrecord_mi_created_by_sm', getByHeader(row, headerMap, 'custrecord_mi_created_by_sm'));

        setTextOrValue(rec, 'custrecord_memo', getByHeader(row, headerMap, 'custrecord_memo'));

        var recId = nlapiSubmitRecord(rec, true, true);

        created++;
        nlapiLogExecution('AUDIT', 'Record Created', 'Row: ' + r + ', ID: ' + recId);

      } catch (e) {
        failed++;
        nlapiLogExecution('ERROR', 'Row Failed ' + r, e.toString());
      }
    }

    nlapiLogExecution('AUDIT', 'Process Complete', 'Created: ' + created + ', Failed: ' + failed);

  } catch (e) {
    nlapiLogExecution('ERROR', 'Critical Error', e.toString());
  }
}

function buildHeaderMap(headers) {
  var map = {};

  for (var i = 0; i < headers.length; i++) {
    var h = normalize(headers[i]);

    map[h] = i;

    if (
      h === 'custrecord ds start date' ||
      h === 'accepted start date'
    ) {
      map['custrecord_ds_start_date'] = i;
    }

    if (
      h === 'custrecord ds end date' ||
      h === 'end date'
    ) {
      map['custrecord_ds_end_date'] = i;
    }

    if (
      h === 'custrecord ds price currency' ||
      h === 'currency' ||
      h === 'price currency'
    ) {
      map['custrecord_ds_price_currency'] = i;
    }

    if (
      h === 'custrecord ds old price' ||
      h === 'old price' ||
      h === 'oldprice'
    ) {
      map['custrecord_ds_old_price'] = i;
    }

    if (
      h === 'custrecord ds new price' ||
      h === 'accepted price' ||
      h === 'new price'
    ) {
      map['custrecord_ds_new_price'] = i;
    }

    if (
      h === 'custrecord ds price level' ||
      h === 'price level' ||
      h === 'price level internal id' ||
      h === 'price level id'
    ) {
      map['custrecord_ds_price_level'] = i;
    }

    if (
      h === 'custrecord ds price level item' ||
      h === 'price level item' ||
      h === 'item' ||
      h === 'item name' ||
      h === 'item internal id' ||
      h === 'item id'
    ) {
      map['custrecord_ds_price_level_item'] = i;
    }

    if (
      h === 'custrecord160'
    ) {
      map['custrecord160'] = i;
    }

    if (
      h === 'custrecordis lot' ||
      h === 'is lot' ||
      h === 'lot'
    ) {
      map['custrecordis_lot'] = i;
    }

    // FIXED: supports internal ID column headers for custrecord_new_price_level
    if (
      h === 'custrecord new price level' ||
      h === 'new price level' ||
      h === 'new price level internal id' ||
      h === 'new price level id' ||
      h === 'price level internal id' ||
      h === 'price level id'
    ) {
      map['custrecord_new_price_level'] = i;
    }

    if (
      h === 'custrecord mi customers' ||
      h === 'customer' ||
      h === 'customers' ||
      h === 'customer internal id' ||
      h === 'customer id'
    ) {
      map['custrecord_mi_customers'] = i;
    }

    if (
      h === 'custrecord mi created by sm' ||
      h === 'created by sm'
    ) {
      map['custrecord_mi_created_by_sm'] = i;
    }

    if (
      h === 'custrecord memo' ||
      h === 'memo'
    ) {
      map['custrecord_memo'] = i;
    }
  }

  nlapiLogExecution('DEBUG', 'Header Map', JSON.stringify(map));

  return map;
}

function getByHeader(row, headerMap, fieldId) {
  var idx = headerMap[fieldId];

  if (idx === undefined || idx === null) {
    nlapiLogExecution('DEBUG', 'Missing Header', fieldId);
    return '';
  }

  var value = row[idx] || '';

  nlapiLogExecution('DEBUG', 'CSV Value', fieldId + ' = ' + value);

  return value;
}

function setTextOrValue(rec, fieldId, value) {
  if (!isEmpty(value)) {
    rec.setFieldValue(fieldId, String(value).trim());
  }
}

function setDateValue(rec, fieldId, value) {
  if (!isEmpty(value)) {
    rec.setFieldValue(fieldId, value);
  }
}

function parseDateValue(value) {
  if (isEmpty(value)) return '';

  value = String(value).trim();

  var parts;
  var dt;

  // DD/MM/YYYY
  if (value.indexOf('/') !== -1) {
    parts = value.split('/');

    if (parts.length === 3) {
      dt = nlapiStringToDate(parts[1] + '/' + parts[0] + '/' + parts[2]);
      return nlapiDateToString(dt);
    }
  }

  // YYYY-MM-DD
  if (value.indexOf('-') !== -1) {
    parts = value.split('-');

    if (parts.length === 3 && parts[0].length === 4) {
      dt = nlapiStringToDate(parts[1] + '/' + parts[2] + '/' + parts[0]);
      return nlapiDateToString(dt);
    }
  }

  return value;
}

function setCheckboxValue(rec, fieldId, value) {
  if (isEmpty(value)) return;

  var v = String(value).trim().toLowerCase();

  if (v === 't' || v === 'true' || v === 'yes' || v === 'y' || v === '1') {
    rec.setFieldValue(fieldId, 'T');
  } else {
    rec.setFieldValue(fieldId, 'F');
  }
}

function setSelectValue(rec, fieldId, value, recordType) {
  if (isEmpty(value)) return;

  var finalId = findRecordId(recordType, value);

  if (!finalId) {
    throw 'Could not find value for field ' + fieldId + ': ' + value;
  }

  rec.setFieldValue(fieldId, finalId);
}

function setCustomerValue(rec, fieldId, value) {
  if (isEmpty(value)) return;

  var parts = String(value).split(/[;|]/);
  var ids = [];

  for (var i = 0; i < parts.length; i++) {
    if (!isEmpty(parts[i])) {
      var id = findRecordId('customer', parts[i]);

      if (!id) {
        throw 'Could not find customer: ' + parts[i];
      }

      ids.push(id);
    }
  }

  if (ids.length === 1) {
    rec.setFieldValue(fieldId, ids[0]);
  } else if (ids.length > 1) {
    rec.setFieldValues(fieldId, ids);
  }
}

function findRecordId(recordType, value) {
  value = String(value || '').trim();

  if (isEmpty(value)) return '';

  // IMPORTANT:
  // If CSV value is internal ID, populate directly.
  if (/^\d+$/.test(value)) {
    return value;
  }

  var filtersList = [];

  if (recordType === 'item') {
    filtersList = [
      [new nlobjSearchFilter('itemid', null, 'is', value)],
      [new nlobjSearchFilter('name', null, 'is', value)],
      [new nlobjSearchFilter('displayname', null, 'is', value)]
    ];
  } else if (recordType === 'customer') {
    filtersList = [
      [new nlobjSearchFilter('entityid', null, 'is', value)],
      [new nlobjSearchFilter('altname', null, 'is', value)],
      [new nlobjSearchFilter('companyname', null, 'is', value)]
    ];
  } else if (recordType === 'currency') {
    filtersList = [
      [new nlobjSearchFilter('name', null, 'is', value)],
      [new nlobjSearchFilter('symbol', null, 'is', value)]
    ];
  } else if (recordType === 'pricelevel') {
    filtersList = [
      [new nlobjSearchFilter('name', null, 'is', value)]
    ];
  } else {
    filtersList = [
      [new nlobjSearchFilter('name', null, 'is', value)]
    ];
  }

  for (var i = 0; i < filtersList.length; i++) {
    try {
      var results = nlapiSearchRecord(
        recordType,
        null,
        filtersList[i],
        [new nlobjSearchColumn('internalid')]
      );

      if (results && results.length > 0) {
        return results[0].getId();
      }
    } catch (e) {
      nlapiLogExecution('DEBUG', 'Lookup Failed', recordType + ' / ' + value + ' / ' + e.toString());
    }
  }

  return '';
}

function parseCSV(content) {
  var lines = String(content).split(/\r\n|\n/);
  var result = [];

  for (var i = 0; i < lines.length; i++) {
    if (!lines[i] || String(lines[i]).trim() === '') continue;
    result.push(parseCSVLine(lines[i]));
  }

  return result;
}

function parseCSVLine(line) {
  var out = [];
  var cur = '';
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var ch = line.charAt(i);

    if (ch === '"') {
      if (inQuotes && line.charAt(i + 1) === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      out.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }

  out.push(cur);

  for (var j = 0; j < out.length; j++) {
    out[j] = out[j] ? String(out[j]).trim() : '';
  }

  return out;
}

function isBlankRow(row) {
  for (var i = 0; i < row.length; i++) {
    if (!isEmpty(row[i])) return false;
  }

  return true;
}

function cleanNumber(v) {
  if (isEmpty(v)) return '';
  return String(v).replace(/,/g, '').replace(/\s+/g, '');
}

function normalize(v) {
  return String(v || '')
    .replace(/^\uFEFF/, '')
    .trim()
    .toLowerCase()
    .replace(/_/g, ' ')
    .replace(/-/g, ' ')
    .replace(/\s+/g, ' ');
}

function isEmpty(v) {
  return v === null || v === undefined || String(v).trim() === '';
}

function looksBase64(s) {
  if (!s) return false;

  s = String(s).replace(/\s+/g, '');

  if (s.length < 50) return false;

  return /^[A-Za-z0-9+/=]+$/.test(s);
}

function base64Decode(b64) {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  var str = String(b64).replace(/[\r\n\s]/g, '');
  var output = '';
  var buffer = 0;
  var bits = 0;

  for (var i = 0; i < str.length; i++) {
    var c = str.charAt(i);

    if (c === '=') break;

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