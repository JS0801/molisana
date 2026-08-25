/**
 * Email Capture Plug-in (SuiteScript 1.0)
 *
 * Existing behavior:
 *   1. Matches the subject line against SUBJECT_IMPORT_MAP
 *   2. Saves the CSV attachment to the File Cabinet
 *   3. Creates the LCL Vendor Credit/Credit Memo first when applicable
 *   4. Schedules the 2.x CSV import trigger script
 *
 * Add-on behavior:
 *   - LCL Deductions_timestamp_amount creates a Vendor Credit
 *   - LCL Remittances_timestamp_amount creates a Credit Memo
 *
 * Expected LCL subject examples:
 *   LCL Deductions_2026-08-10 17:25:38Z_1326.33
 *   LCL Remittances_2026-08-10 17:25:38Z_1326.33
 *
 * Expected LCL attachment filename examples:
 *   Deduction_2002161177.csv
 *   Remittance_2002161177.csv
 *   Deduction_2002161177 (2).csv
 *   Remittance_2003851933 (1).csv
 *
 * The transaction number is the numeric filename portion, e.g. 2002161177.
 */

// ------------------------------------------------------------------
// CONFIG
// ------------------------------------------------------------------
var SUBJECT_IMPORT_MAP = [
    { keyword: 'Deduction',  mappingId: '157', description: 'Deduction CSV import' },
    { keyword: 'Remittance', mappingId: '156', description: 'Remittance CSV import' }
];

var TARGET_FOLDER_ID = 475348;

var SCHEDULED_SCRIPT_ID = 'customscript_mi_import_bills_and_payment';
var SCHEDULED_DEPLOYMENT_ID = 'customdeploy_mi_import_bills_and_payment';

var LCL_DEDUCTION_TYPE = 'deduction';
var LCL_REMITTANCE_TYPE = 'remittance';
var LCL_CUSTOMER_KEY = 'lcl';
var METRO_CUSTOMER_KEY = 'metro';
var ALLOWED_SENDER_DOMAIN = 'molisana.com';

var LCL_TRANSACTION_CONFIG = {};

function getTransactionConfigKey(customerKey, transactionType) {
    return customerKey + '_' + transactionType;
}

function normalizeExternalIdPart(value) {
    return trim(value)
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '_')
        .replace(/^_+|_+$/g, '');
}

function buildEmailCaptureExternalId(transactionName, remittanceNumber) {
    return normalizeExternalIdPart(transactionName) + '_' + normalizeExternalIdPart(remittanceNumber);
}

function getRefundCustomerKey(customerId) {
    if (String(customerId) === '30') return 'lcl';
    if (String(customerId) === '6327') return 'metro';
    return 'customer_' + customerId;
}

function findExistingEmailCaptureTransaction(recordType, externalId) {
    var fields = ['externalidstring', 'externalid'];

    for (var i = 0; i < fields.length; i++) {
        try {
            var results = nlapiSearchRecord(recordType, null, [
                new nlobjSearchFilter(fields[i], null, 'is', externalId)
            ], [
                new nlobjSearchColumn('internalid')
            ]);

            if (results && results.length) {
                return results[0].getId();
            }
        } catch (e) {
            debugLog('External ID lookup failed', 'recordType=' + recordType + ' externalId=' + externalId + ' field=' + fields[i] + ' error=' + getErrorDetails(e));
        }
    }

    return '';
}

function getLclTransactionConfig(customerKey, transactionType) {
    return LCL_TRANSACTION_CONFIG[getTransactionConfigKey(customerKey, transactionType)];
}

LCL_TRANSACTION_CONFIG[getTransactionConfigKey(LCL_CUSTOMER_KEY, LCL_DEDUCTION_TYPE)] = {
    label: 'LCL deduction vendor credit',
    recordType: 'vendorcredit',
    entityId: '11437',
    accountId: '2059',
    itemId: '6733',
    fileType: 'Deduction',
    dateColumnName: 'Invoice Date',
    referenceFieldId: 'custbody_note_to_vendor',
    setLineLocation: true,
    setLineDescription: false
};

LCL_TRANSACTION_CONFIG[getTransactionConfigKey(LCL_CUSTOMER_KEY, LCL_REMITTANCE_TYPE)] = {
    label: 'LCL remittance credit memo',
    recordType: 'creditmemo',
    entityId: '30',
    accountId: '119',
    itemId: '6808',
    fileType: 'Remittance',
    dateColumnName: 'Payment Date',
    referenceFieldId: 'custbody_2663_reference_num',
    setLineLocation: false,
    setLineDescription: true
};

LCL_TRANSACTION_CONFIG[getTransactionConfigKey(METRO_CUSTOMER_KEY, LCL_DEDUCTION_TYPE)] = {
    label: 'Metro deduction vendor credit',
    recordType: 'vendorcredit',
    entityId: '442',
    accountId: '2059',
    itemId: '6733',
    fileType: 'Deduction',
    dateColumnName: 'Invoice Date',
    referenceFieldId: 'custbody_note_to_vendor',
    setLineLocation: true,
    setLineDescription: false
};

LCL_TRANSACTION_CONFIG[getTransactionConfigKey(METRO_CUSTOMER_KEY, LCL_REMITTANCE_TYPE)] = {
    label: 'Metro remittance credit memo',
    recordType: 'creditmemo',
    entityId: '6327',
    accountId: '119',
    itemId: '6113',
    fileType: 'Remittance',
    dateColumnName: 'Remittance Date',
    referenceFieldId: 'custbody_2663_reference_num',
    setLineLocation: false,
    setLineDescription: true
};

var LCL_CURRENCY_ID = '1';
var LCL_LOCATION_ID = '315';
var LCL_CLASS_ID = '319';
var LCL_BRAND_ID = '227';
var LCL_TAX_CODE_ID = '16';
var LCL_UNITS_ID = '35';
var CUSTOMER_REFUND_ACCOUNT_ID = '2058';
var METRO_CREDIT_COLUMN_NAMES = {
    creditNum: 'Credit Num',
    creditInternalId: 'Credit Internal ID',
    creditGross: 'Credit Gross',
    creditDiscount: 'Credit Discount',
    creditNet: 'Credit Net'
};
var METRO_UPDATED_FILE_PREFIX = 'update_file_';

// Keep DEBUG on while testing; change to false after deployment is stable.
var LCL_DEBUG_LOGS = true;

// ------------------------------------------------------------------
// EXISTING CSV IMPORT HANDOFF
// ------------------------------------------------------------------

/**
 * Finds the first matching import config for a subject line.
 * @param {string} subject
 * @returns {Object|null}
 */
function resolveImportConfig(subject) {
    if (!subject) return null;
    var subjectLower = subject.toLowerCase();

    for (var i = 0; i < SUBJECT_IMPORT_MAP.length; i++) {
        if (subjectLower.indexOf(SUBJECT_IMPORT_MAP[i].keyword.toLowerCase()) !== -1) {
            return SUBJECT_IMPORT_MAP[i];
        }
    }
    return null;
}

/**
 * Pulls the first .csv attachment off the inbound email.
 * @param {nlobjEmail} message
 * @returns {nlobjFile|null}
 */
function getCsvAttachment(message) {
    var attachments = message.getAttachments() || [];
    debugLog('Attachment scan started', 'attachmentCount=' + attachments.length);

    for (var i = 0; i < attachments.length; i++) {
        var name = attachments[i].getName() || '';
        debugLog('Attachment found', 'index=' + i + ' name=' + name);
        if (name.toLowerCase().indexOf('.csv') !== -1) {
            debugLog('CSV attachment selected', 'index=' + i + ' name=' + name);
            return attachments[i];
        }
    }
    return null;
}

/**
 * Saves the attachment to the File Cabinet, keeping its original name.
 * @param {nlobjFile} csvFile
 * @returns {string} file internal id
 */
function saveAttachment(csvFile) {
    debugLog('Saving CSV attachment', 'targetFolderId=' + TARGET_FOLDER_ID + ' originalName=' + csvFile.getName());
    csvFile.setFolder(TARGET_FOLDER_ID);
    var fileId = nlapiSubmitFile(csvFile);

    nlapiLogExecution('AUDIT', 'CSV attachment saved', 'fileId=' + fileId + ' name=' + csvFile.getName());
    return fileId;
}

/**
 * Hands off to the 2.x scheduled script that will actually run
 * N/task CSV_IMPORT.
 * @param {string} fileId
 * @param {string} mappingId
 */
function triggerCsvImport(fileId, mappingId) {
    debugLog('Scheduling CSV import trigger', 'scriptId=' + SCHEDULED_SCRIPT_ID + ' deploymentId=' + SCHEDULED_DEPLOYMENT_ID + ' fileId=' + fileId + ' mappingId=' + mappingId);
    var status = nlapiScheduleScript(SCHEDULED_SCRIPT_ID, SCHEDULED_DEPLOYMENT_ID, {
        custscript_import_file_id: fileId,
        custscript_import_mapping_id: mappingId
    });

    nlapiLogExecution('AUDIT', 'CSV import scheduled', 'fileId=' + fileId + ' mappingId=' + mappingId + ' status=' + status);
}

// ------------------------------------------------------------------
// LCL SUBJECT PARSING
// ------------------------------------------------------------------

function trim(value) {
    return String(value || '').replace(/^\s+|\s+$/g, '');
}

function parseAmount(amountText) {
    var normalized = trim(amountText).replace(/,/g, '');

    if (!/^\d+(\.\d{1,2})?$/.test(normalized)) {
        return null;
    }

    var amount = parseFloat(normalized);
    if (!(amount > 0)) {
        return null;
    }

    return amount.toFixed(2);
}

function isZeroOrEmptyAmount(amountText) {
    var normalized = trim(amountText).replace(/,/g, '');

    if (normalized === '') {
        return true;
    }

    return parseFloat(normalized) === 0;
}

function parseUtcTimestamp(timestampText) {
    var match = /^(\d{4})-(\d{2})-(\d{2}) ([0-2]\d):([0-5]\d):([0-5]\d)Z$/.exec(trim(timestampText));
    if (!match) {
        return null;
    }

    var year = parseInt(match[1], 10);
    var month = parseInt(match[2], 10);
    var day = parseInt(match[3], 10);
    var hour = parseInt(match[4], 10);
    var minute = parseInt(match[5], 10);
    var second = parseInt(match[6], 10);

    if (month < 1 || month > 12 || hour > 23) {
        return null;
    }

    var date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));

    if (date.getUTCFullYear() !== year ||
            date.getUTCMonth() !== month - 1 ||
            date.getUTCDate() !== day ||
            date.getUTCHours() !== hour ||
            date.getUTCMinutes() !== minute ||
            date.getUTCSeconds() !== second) {
        return null;
    }

    return date;
}

/**
 * Parses only the new LCL subject patterns. Other Deduction/Remittance
 * subjects still run through the existing CSV import path only.
 * @param {string} subject
 * @returns {Object|null}
 */
function parseLclSubject(subject) {
    if (!isLclTransactionSubject(subject)) {
        return null;
    }

    var match = /^\s*((LCL|Metro)\s+(Deductions|Remittances))\s*(?:_|-\s*)([^_]+?)(?:_([^_]*))?(?:_([^_]*))?\s*\.?\s*$/i.exec(String(subject));
    if (!match) {
        return { error: 'Expected subject like LCL Deductions - timestamp_amount or Metro Deductions - timestamp_amount' };
    }

    var customerText = trim(match[2]);
    var typeText = trim(match[3]);
    var timestampText = trim(match[4]).replace(/\.$/, '');
    var amountText = trim(match[5] || '');
    var secondAmountText = trim(match[6] || '');
    var customerKey = null;
    var customerName = null;
    var transactionType = null;

    if (/^LCL$/i.test(customerText)) {
        customerKey = LCL_CUSTOMER_KEY;
        customerName = 'LCL';
    } else if (/^Metro$/i.test(customerText)) {
        customerKey = METRO_CUSTOMER_KEY;
        customerName = 'Metro';
    } else {
        return { error: 'Unsupported customer in subject: ' + customerText };
    }

    if (/^Deductions$/i.test(typeText)) {
        transactionType = LCL_DEDUCTION_TYPE;
    } else if (/^Remittances$/i.test(typeText)) {
        transactionType = LCL_REMITTANCE_TYPE;
    } else {
        return { error: 'Unsupported transaction type in subject: ' + typeText };
    }

    var timestampDate = parseUtcTimestamp(timestampText);
    if (!timestampDate) {
        return { error: 'Invalid timestamp. Expected YYYY-MM-DD HH:MM:SSZ, received: ' + timestampText };
    }

var amount = '0.00';
var skipTransaction = false;
var skipReason = '';

if (isZeroOrEmptyAmount(amountText)) {
    skipTransaction = true;
    skipReason = 'First amount is zero or blank';
} else {
    amount = parseAmount(amountText);
    if (!amount) {
        return { error: 'Invalid amount. Expected a positive decimal amount, received: ' + amountText };
    }
}

var secondAmount = '0.00';
var skipSecondBillCredit = true;
var secondSkipReason = 'Second amount is zero or blank';

if (!isZeroOrEmptyAmount(secondAmountText)) {
    secondAmount = parseAmount(secondAmountText);
    if (!secondAmount) {
        return { error: 'Invalid second amount. Expected a positive decimal amount, received: ' + secondAmountText };
    }

    skipSecondBillCredit = false;
    secondSkipReason = '';
}

return {
    customerKey: customerKey,
    customerName: customerName,
    transactionType: transactionType,
    timestampText: timestampText,
    timestampDate: timestampDate,
    amount: amount,
    skipTransaction: skipTransaction,
    skipReason: skipReason,
    secondAmount: secondAmount,
    skipSecondBillCredit: skipSecondBillCredit,
    secondSkipReason: secondSkipReason
};
}
function createPaybackVendorCredit(lclSubject, lclFile, fileName, transactionDate) {
    var config = getLclTransactionConfig(lclSubject.customerKey, LCL_DEDUCTION_TYPE);
    var documentNumber = lclFile.documentNumber;


    var paybackDocumentNumber = documentNumber + '_payback';


  var externalId = buildEmailCaptureExternalId(lclSubject.customerKey + '_payback', documentNumber);
var existingRecordId = findExistingEmailCaptureTransaction('vendorcredit', externalId);

if (existingRecordId) {
    nlapiLogExecution('AUDIT', 'Payback Vendor Credit already exists', 'externalId=' + externalId + ' existingId=' + existingRecordId + ' tranid=' + paybackDocumentNumber);

    return {
        id: existingRecordId,
        recordType: 'vendorcredit',
        tranid: paybackDocumentNumber
    };
}

    var record = nlapiCreateRecord('vendorcredit', { recordmode: 'dynamic' });
    var memo = 'EFT ' + paybackDocumentNumber;
    setBodyField(record, 'externalid', externalId);
    setBodyField(record, 'entity', config.entityId);
    setBodyField(record, 'trandate', transactionDate);
    setBodyField(record, 'tranid', paybackDocumentNumber);
    setBodyField(record, 'memo', memo);
    setBodyField(record, 'custbody_created_from_email_capture', 'T');
    setBodyField(record, 'currency', LCL_CURRENCY_ID);
    setBodyField(record, 'location', LCL_LOCATION_ID);
    setBodyField(record, 'account', config.accountId);
    setBodyField(record, 'custbody_report_timestamp', getNetSuiteDateTimeValue(lclSubject.timestampDate));
    setBodyField(record, config.referenceFieldId, paybackDocumentNumber);

    addLclItemLine(record, config, lclSubject.secondAmount, paybackDocumentNumber, fileName);

    var recordId = nlapiSubmitRecord(record, true, true);
    nlapiLogExecution('AUDIT', 'Payback Vendor Credit created', 'id=' + recordId + ' tranid=' + paybackDocumentNumber + ' amount=' + lclSubject.secondAmount);

    return {
        id: recordId,
        recordType: 'vendorcredit',
        tranid: paybackDocumentNumber
    };
}
function isLclTransactionSubject(subject) {
    return /^\s*(LCL|Metro)\s+(Deductions|Remittances)\s*(?:_|-\s*)/i.test(subject || '');
}

function parseLclCsvFileName(fileName) {
    var match = /^(?:(LCL|Metro)\s+)?(Deduction|Remittance)_([0-9]+)(?:\s+\([0-9]+\))?\.csv$/i.exec(trim(fileName));
    if (!match) {
        return { error: 'Invalid CSV filename. Expected optional customer prefix plus Deduction_number.csv or Remittance_number.csv, received: ' + fileName };
    }

    var customerText = trim(match[1] || '');
    var fileType = match[2];

    return {
        customerKey: /^LCL$/i.test(customerText) ? LCL_CUSTOMER_KEY : /^Metro$/i.test(customerText) ? METRO_CUSTOMER_KEY : '',
        customerName: customerText,
        transactionType: /^Deduction$/i.test(fileType) ? LCL_DEDUCTION_TYPE : LCL_REMITTANCE_TYPE,
        fileType: fileType,
        documentNumber: match[3]
    };
}

function normalizeHeaderName(value) {
    return trim(value).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function findCsvColumnIndex(headers, possibleNames) {
    for (var i = 0; i < possibleNames.length; i++) {
        var expected = normalizeHeaderName(possibleNames[i]);
        for (var j = 0; j < headers.length; j++) {
            if (normalizeHeaderName(headers[j]) === expected) {
                return j;
            }
        }
    }

    return -1;
}

function parseCsvAmountValue(value) {
    var text = trim(value);
    if (!text) {
        return 0;
    }

    var negative = /^\(.*\)$/.test(text) || text.charAt(0) === '-';
    var normalized = text.replace(/[,$()\s]/g, '').replace(/^-/, '');
    if (!/^\d+(\.\d+)?$/.test(normalized)) {
        return 0;
    }

    var amount = parseFloat(normalized);
    return negative ? -amount : amount;
}

function formatCsvAmount(value) {
    var amount = Math.abs(parseCsvAmountValue(value));
    return String(amount);
}

function getCsvValue(values, index) {
    if (index < 0 || index >= values.length) {
        return '';
    }

    return trim(values[index]);
}

function getFirstCsvValue(values, indexes) {
    for (var i = 0; i < indexes.length; i++) {
        var value = getCsvValue(values, indexes[i]);
        if (value) {
            return value;
        }
    }

    return '';
}

function getCreditInternalIdColumnIndexes(headers) {
    var indexes = [];
    var exactNames = [
        'Credit Internal ID',
        'Credit Memo Internal ID',
        'Credit Memo ID',
        'Credit ID',
        'Credit Document Internal ID',
        'Invoice Internal ID'
    ];

    for (var i = 0; i < exactNames.length; i++) {
        var index = findCsvColumnIndex(headers, [exactNames[i]]);
        if (index !== -1) {
            indexes.push(index);
        }
    }

    return indexes;
}

function getCreditTranIdColumnIndexes(headers) {
    var indexes = [];
    var exactNames = [
        'Credit Num',
        'Credit Number',
        'Credit #',
        'Credit Memo #',
        'Credit Reference',
        'Credit Memo Reference',
        'Invoice Number',
        'Invoice #',
        'Invoice Reference'
    ];

    for (var i = 0; i < exactNames.length; i++) {
        var index = findCsvColumnIndex(headers, [exactNames[i]]);
        if (index !== -1) {
            indexes.push(index);
        }
    }

    return indexes;
}

function getPayloadHeaderIndexes(headers) {
    return {
        paymentExternalId: findCsvColumnIndex(headers, ['Payment External ID']),
        paymentNumber: findCsvColumnIndex(headers, ['Payment #', 'Remittance Number']),
        paymentDate: findCsvColumnIndex(headers, ['Payment Date', 'Remittance Date']),
        parentId: findCsvColumnIndex(headers, ['Parent ID'])
    };
}

function ensureCsvColumn(headers, columnName) {
    var index = findCsvColumnIndex(headers, [columnName]);
    if (index === -1) {
        index = headers.length;
        headers.push(columnName);
    }

    return index;
}

function ensureCsvRowLength(row, length) {
    while (row.length < length) {
        row.push('');
    }
}

function getMetroCreditColumnIndexes(headers) {
    return {
        creditNum: ensureCsvColumn(headers, METRO_CREDIT_COLUMN_NAMES.creditNum),
        creditInternalId: ensureCsvColumn(headers, METRO_CREDIT_COLUMN_NAMES.creditInternalId),
        creditGross: ensureCsvColumn(headers, METRO_CREDIT_COLUMN_NAMES.creditGross),
        creditDiscount: ensureCsvColumn(headers, METRO_CREDIT_COLUMN_NAMES.creditDiscount),
        creditNet: ensureCsvColumn(headers, METRO_CREDIT_COLUMN_NAMES.creditNet)
    };
}

function clearMetroCreditColumns(row, creditColumnIndexes) {
    row[creditColumnIndexes.creditNum] = '';
    row[creditColumnIndexes.creditInternalId] = '';
    row[creditColumnIndexes.creditGross] = '';
    row[creditColumnIndexes.creditDiscount] = '';
    row[creditColumnIndexes.creditNet] = '';
}

function setMetroCreditColumns(row, creditColumnIndexes, credit) {
    row[creditColumnIndexes.creditNum] = credit.creditTranId;
    row[creditColumnIndexes.creditInternalId] = credit.creditInternalId;
    row[creditColumnIndexes.creditGross] = credit.creditGross;
    row[creditColumnIndexes.creditDiscount] = credit.creditDiscount;
    row[creditColumnIndexes.creditNet] = credit.creditNet;
}

function copyCsvValueByColumnName(headers, sourceRow, targetRow, columnName) {
    var index = findCsvColumnIndex(headers, [columnName]);
    if (index !== -1) {
        targetRow[index] = getCsvValue(sourceRow, index);
    }
}

function createMetroCreditOnlyRow(headers, sourceRow, creditColumnIndexes) {
    var row = [];
    ensureCsvRowLength(row, headers.length);

    copyCsvValueByColumnName(headers, sourceRow, row, 'Customer');
    copyCsvValueByColumnName(headers, sourceRow, row, 'Payment External ID');
    copyCsvValueByColumnName(headers, sourceRow, row, 'Payment Date');
    copyCsvValueByColumnName(headers, sourceRow, row, 'Remittance Date');
    copyCsvValueByColumnName(headers, sourceRow, row, 'Payment #');
    copyCsvValueByColumnName(headers, sourceRow, row, 'Remittance Number');
    copyCsvValueByColumnName(headers, sourceRow, row, 'Status');
    copyCsvValueByColumnName(headers, sourceRow, row, 'Memo');
    copyCsvValueByColumnName(headers, sourceRow, row, 'Parent ID');
    clearMetroCreditColumns(row, creditColumnIndexes);

    return row;
}

function applyMetroCreditColumns(headers, keptRows, firstDataRow, creditColumnIndexes, credits) {
    for (var i = 0; i < credits.length; i++) {
        var targetRow = keptRows[i];
        if (!targetRow) {
            targetRow = createMetroCreditOnlyRow(headers, firstDataRow, creditColumnIndexes);
            keptRows.push(targetRow);
        }

        ensureCsvRowLength(targetRow, headers.length);
        setMetroCreditColumns(targetRow, creditColumnIndexes, credits[i]);
    }
}

function csvEscape(value) {
    var text = value === null || value === undefined ? '' : String(value);
    if (/[",\r\n]/.test(text)) {
        return '"' + text.replace(/"/g, '""') + '"';
    }

    return text;
}

function buildCsvContent(headers, dataRows) {
    var lines = [];
    lines.push(headers.map(csvEscape).join(','));

    for (var i = 0; i < dataRows.length; i++) {
        var row = dataRows[i];
        var output = [];
        for (var col = 0; col < headers.length; col++) {
            output.push(csvEscape(row[col] || ''));
        }
        lines.push(output.join(','));
    }

    return lines.join('\r\n');
}

function buildUpdatedFileName(fileName) {
    return METRO_UPDATED_FILE_PREFIX + trim(fileName);
}

function formatCsvSumAmount(value) {
    var amount = Math.round(parseFloat(value || 0) * 100000000) / 100000000;
    if (Math.abs(amount) < 0.000000001) {
        amount = 0;
    }

    var text = String(amount);
    if (text.indexOf('e') !== -1 || text.indexOf('E') !== -1) {
        text = amount.toFixed(8).replace(/0+$/g, '').replace(/\.$/, '');
    }

    return text;
}

function analyzeLclRemittanceRows(csvFile) {
    var contents = getCsvFileContents(csvFile);
    if (!contents) {
        return { error: 'CSV attachment contents are empty or unavailable' };
    }

    var rows = contents.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    if (!rows.length || !trim(rows[0])) {
        return { error: 'CSV header row is empty' };
    }

    var headers = parseCsvLine(rows[0]);
    var invoiceInternalIdIndex = findCsvColumnIndex(headers, ['Invoice Internal ID']);
    var grossIndex = findCsvColumnIndex(headers, ['Gross']);
    var discountIndex = findCsvColumnIndex(headers, ['Discount']);
    var netIndex = findCsvColumnIndex(headers, ['Net']);
    var missingColumns = [];

    if (invoiceInternalIdIndex === -1) missingColumns.push('Invoice Internal ID');
    if (grossIndex === -1) missingColumns.push('Gross');
    if (discountIndex === -1) missingColumns.push('Discount');
    if (netIndex === -1) missingColumns.push('Net');

    if (missingColumns.length) {
        return { error: 'LCL remittance combine required columns missing: ' + missingColumns.join(', ') };
    }

    var groupedByInvoiceInternalId = {};
    var groups = [];
    var sourceRowCount = 0;
    var combinedSourceRowCount = 0;

    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
        if (!trim(rows[rowIndex])) {
            continue;
        }

        var values = parseCsvLine(rows[rowIndex]);
        ensureCsvRowLength(values, headers.length);

        var invoiceInternalId = getCsvValue(values, invoiceInternalIdIndex);
        var groupKey = invoiceInternalId || ('__blank_invoice_internal_id_row_' + rowIndex);
        var group = groupedByInvoiceInternalId[groupKey];
        sourceRowCount++;

        if (!group) {
            group = {
                row: values,
                invoiceInternalId: invoiceInternalId,
                rowCount: 0,
                grossTotal: 0,
                discountTotal: 0,
                netTotal: 0
            };
            groupedByInvoiceInternalId[groupKey] = group;
            groups.push(group);
        } else {
            combinedSourceRowCount++;
        }

        group.rowCount++;
        group.grossTotal += parseCsvAmountValue(getCsvValue(values, grossIndex));
        group.discountTotal += parseCsvAmountValue(getCsvValue(values, discountIndex));
        group.netTotal += parseCsvAmountValue(getCsvValue(values, netIndex));
    }

    var outputRows = [];
    var combinedGroupCount = 0;
    var removedNegativeGroupCount = 0;
    var removedNegativeSourceRowCount = 0;

    for (var groupIndex = 0; groupIndex < groups.length; groupIndex++) {
        var currentGroup = groups[groupIndex];
        if (currentGroup.rowCount > 1) {
            combinedGroupCount++;
            currentGroup.row[grossIndex] = formatCsvSumAmount(currentGroup.grossTotal);
            currentGroup.row[discountIndex] = formatCsvSumAmount(currentGroup.discountTotal);
            currentGroup.row[netIndex] = formatCsvSumAmount(currentGroup.netTotal);
        }

        if (currentGroup.netTotal < -0.000000001) {
            removedNegativeGroupCount++;
            removedNegativeSourceRowCount += currentGroup.rowCount;
            continue;
        }

        outputRows.push(currentGroup.row);
    }

    if (!outputRows.length && sourceRowCount) {
        return { error: 'All LCL remittance rows have negative Net after combining. Cannot create payment import file with no payment rows.' };
    }

    if (!combinedSourceRowCount && !removedNegativeGroupCount) {
        debugLog('LCL remittance combine skipped', 'No duplicate Invoice Internal ID rows or negative Net rows found. fileName=' + csvFile.getName());
        return { updated: false };
    }

    debugLog('LCL remittance rows combined',
            'fileName=' + csvFile.getName() +
            ' sourceRowCount=' + sourceRowCount +
            ' outputRowCount=' + outputRows.length +
            ' combinedGroupCount=' + combinedGroupCount +
            ' combinedSourceRowCount=' + combinedSourceRowCount +
            ' removedNegativeGroupCount=' + removedNegativeGroupCount +
            ' removedNegativeSourceRowCount=' + removedNegativeSourceRowCount);

    return {
        updated: true,
        combinedGroupCount: combinedGroupCount,
        combinedSourceRowCount: combinedSourceRowCount,
        removedNegativeGroupCount: removedNegativeGroupCount,
        removedNegativeSourceRowCount: removedNegativeSourceRowCount,
        updatedContents: buildCsvContent(headers, outputRows),
        updatedFileName: buildUpdatedFileName(csvFile.getName())
    };
}

function maybeCreateLclRemittanceUpdatedFile(subject, csvFile) {
    var lclSubject = parseLclSubject(subject);
    if (!lclSubject || lclSubject.error ||
            lclSubject.customerKey !== LCL_CUSTOMER_KEY ||
            lclSubject.transactionType !== LCL_REMITTANCE_TYPE) {
        return { updated: false };
    }

    var lclFile = parseLclCsvFileName(csvFile.getName());
    if (lclFile.error) {
        return { error: lclFile.error };
    }

    if (lclFile.transactionType !== lclSubject.transactionType) {
        return { error: 'Subject/file type mismatch for LCL remittance combine. subject=' + subject + ' fileName=' + csvFile.getName() };
    }

    if (lclFile.customerKey && lclFile.customerKey !== lclSubject.customerKey) {
        return { error: 'Subject/file customer mismatch for LCL remittance combine. subject=' + subject + ' fileName=' + csvFile.getName() };
    }

    var analysis = analyzeLclRemittanceRows(csvFile);
    if (analysis.error) {
        return analysis;
    }

    if (!analysis.updated) {
        return { updated: false };
    }

    var updatedFile = nlapiCreateFile(analysis.updatedFileName, 'CSV', analysis.updatedContents);
    updatedFile.setFolder(TARGET_FOLDER_ID);
    var updatedFileId = nlapiSubmitFile(updatedFile);

    nlapiLogExecution('AUDIT', 'LCL remittance updated CSV created',
            'originalFile=' + csvFile.getName() +
            ' updatedFile=' + analysis.updatedFileName +
            ' updatedFileId=' + updatedFileId +
            ' combinedGroupCount=' + analysis.combinedGroupCount +
            ' combinedSourceRowCount=' + analysis.combinedSourceRowCount +
            ' removedNegativeGroupCount=' + analysis.removedNegativeGroupCount +
            ' removedNegativeSourceRowCount=' + analysis.removedNegativeSourceRowCount);

    return {
        updated: true,
        fileId: updatedFileId,
        fileName: analysis.updatedFileName,
        combinedGroupCount: analysis.combinedGroupCount,
        combinedSourceRowCount: analysis.combinedSourceRowCount,
        removedNegativeGroupCount: analysis.removedNegativeGroupCount,
        removedNegativeSourceRowCount: analysis.removedNegativeSourceRowCount
    };
}

function analyzeMetroRemittanceCredits(csvFile, lclFile) {
    var contents = getCsvFileContents(csvFile);
    if (!contents) {
        return { error: 'CSV attachment contents are empty or unavailable' };
    }

    var rows = contents.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    if (!rows.length || !trim(rows[0])) {
        return { error: 'CSV header row is empty' };
    }

    var headers = parseCsvLine(rows[0]);
    var netIndex = findCsvColumnIndex(headers, ['Net']);
    if (netIndex === -1) {
        debugLog('Metro credit filter skipped', 'No Net column found. fileName=' + csvFile.getName());
        return { hasCredits: false };
    }

    var grossIndex = findCsvColumnIndex(headers, ['Gross']);
    var discountIndex = findCsvColumnIndex(headers, ['Discount']);
    var creditColumnIndexes = getMetroCreditColumnIndexes(headers);
    var creditInternalIdIndexes = getCreditInternalIdColumnIndexes(headers);
    var creditTranIdIndexes = getCreditTranIdColumnIndexes(headers);
    var keptRows = [];
    var credits = [];
    var firstDataRow = null;

    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
        if (!trim(rows[rowIndex])) {
            continue;
        }

        var values = parseCsvLine(rows[rowIndex]);
        var netAmount = parseCsvAmountValue(getCsvValue(values, netIndex));

        if (netAmount < 0) {
            credits.push({
                rowNumber: rowIndex + 1,
                creditGross: getCsvValue(values, creditColumnIndexes.creditGross) || formatCsvAmount(getCsvValue(values, grossIndex)),
                creditDiscount: getCsvValue(values, creditColumnIndexes.creditDiscount) || formatCsvAmount(getCsvValue(values, discountIndex)),
                creditNet: getCsvValue(values, creditColumnIndexes.creditNet) || formatCsvAmount(netAmount),
                creditInternalId: getFirstCsvValue(values, creditInternalIdIndexes),
                creditTranId: getFirstCsvValue(values, creditTranIdIndexes)
            });
        } else {
            ensureCsvRowLength(values, headers.length);
            clearMetroCreditColumns(values, creditColumnIndexes);
            keptRows.push(values);
            if (!firstDataRow) {
                firstDataRow = values;
            }
        }
    }

    if (!credits.length) {
        debugLog('Metro credit filter skipped', 'Net column found but no negative Net rows. fileName=' + csvFile.getName());
        return { hasCredits: false };
    }

    if (!keptRows.length) {
        return { error: 'All rows have negative Net values. Cannot create payment import file with no payment rows.' };
    }

    applyMetroCreditColumns(headers, keptRows, firstDataRow, creditColumnIndexes, credits);
    debugLog('Metro remittance credit columns added',
            'fileName=' + csvFile.getName() +
            ' creditCount=' + credits.length +
            ' paymentRowCount=' + keptRows.length);

    return {
        hasCredits: true,
        creditCount: credits.length,
        updatedContents: buildCsvContent(headers, keptRows),
        updatedFileName: buildUpdatedFileName(csvFile.getName())
    };
}

function maybeCreateMetroRemittanceUpdatedFile(subject, csvFile) {
    var lclSubject = parseLclSubject(subject);
    if (!lclSubject || lclSubject.error ||
            lclSubject.customerKey !== METRO_CUSTOMER_KEY ||
            lclSubject.transactionType !== LCL_REMITTANCE_TYPE) {
        return { updated: false };
    }

    var lclFile = parseLclCsvFileName(csvFile.getName());
    if (lclFile.error) {
        return { error: lclFile.error };
    }

    if (lclFile.transactionType !== lclSubject.transactionType) {
        return { error: 'Subject/file type mismatch for Metro credit filter. subject=' + subject + ' fileName=' + csvFile.getName() };
    }

    if (lclFile.customerKey && lclFile.customerKey !== lclSubject.customerKey) {
        return { error: 'Subject/file customer mismatch for Metro credit filter. subject=' + subject + ' fileName=' + csvFile.getName() };
    }

    var analysis = analyzeMetroRemittanceCredits(csvFile, lclFile);
    if (analysis.error) {
        return analysis;
    }

    if (!analysis.hasCredits) {
        return { updated: false };
    }

    var updatedFile = nlapiCreateFile(analysis.updatedFileName, 'CSV', analysis.updatedContents);
    updatedFile.setFolder(TARGET_FOLDER_ID);
    var updatedFileId = nlapiSubmitFile(updatedFile);

    nlapiLogExecution('AUDIT', 'Metro remittance updated CSV created',
            'originalFile=' + csvFile.getName() +
            ' updatedFile=' + analysis.updatedFileName +
            ' updatedFileId=' + updatedFileId +
            ' creditCount=' + analysis.creditCount);

    return {
        updated: true,
        fileId: updatedFileId,
        fileName: analysis.updatedFileName,
        creditCount: analysis.creditCount
    };
}

function getCsvFileContents(csvFile) {
    var methodNames = ['getValue', 'getContents', 'getContent'];

    for (var i = 0; i < methodNames.length; i++) {
        try {
            if (csvFile && typeof csvFile[methodNames[i]] === 'function') {
                var value = csvFile[methodNames[i]]();
                if (value !== null && value !== undefined && value !== '') {
                    return String(value);
                }
            }
        } catch (e) {
            debugLog('Unable to read CSV contents', 'method=' + methodNames[i] + ' error=' + getErrorDetails(e));
        }
    }

    return '';
}

function parseCsvLine(line) {
    var values = [];
    var current = '';
    var inQuotes = false;

    for (var i = 0; i < line.length; i++) {
        var character = line.charAt(i);
        var nextCharacter = line.charAt(i + 1);

        if (character === '"') {
            if (inQuotes && nextCharacter === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (character === ',' && !inQuotes) {
            values.push(current);
            current = '';
        } else {
            current += character;
        }
    }

    values.push(current);
    return values;
}

function getCsvColumnValue(csvFile, columnName) {
    var contents = getCsvFileContents(csvFile);
    if (!contents) {
        return { error: 'CSV attachment contents are empty or unavailable' };
    }

    var rows = contents.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    if (!rows.length || !trim(rows[0])) {
        return { error: 'CSV header row is empty' };
    }

    var headers = parseCsvLine(rows[0]);
    var columnIndex = -1;

    for (var i = 0; i < headers.length; i++) {
        if (trim(headers[i]).toLowerCase() === columnName.toLowerCase()) {
            columnIndex = i;
            break;
        }
    }

    if (columnIndex === -1) {
        return { error: 'CSV date column not found: ' + columnName };
    }

    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
        if (!trim(rows[rowIndex])) {
            continue;
        }

        var values = parseCsvLine(rows[rowIndex]);
        var columnValue = trim(values[columnIndex] || '');
        if (columnValue) {
            return {
                value: columnValue,
                rowNumber: rowIndex + 1,
                columnName: columnName
            };
        }
    }

    return { error: 'CSV date column has no value: ' + columnName };
}

function normalizeCsvDateForNetSuite(dateText) {
    var value = trim(dateText);
    var match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(value);
    if (match) {
        return parseInt(match[1], 10) + '/' + parseInt(match[2], 10) + '/' + match[3];
    }

    match = /^(\d{4})-(\d{2})-(\d{2})(?:\s+\d{2}:\d{2}(?::\d{2})?)?$/.exec(value);
    if (match) {
        return parseInt(match[2], 10) + '/' + parseInt(match[3], 10) + '/' + match[1];
    }

    return null;
}

function getCsvTransactionDate(csvFile, config) {
    var dateResult = getCsvColumnValue(csvFile, config.dateColumnName);
    if (dateResult.error) {
        return dateResult;
    }

    var transactionDate = normalizeCsvDateForNetSuite(dateResult.value);
    if (!transactionDate) {
        return { error: 'Invalid CSV transaction date: column=' + config.dateColumnName + ' value=' + dateResult.value };
    }

    debugLog('CSV transaction date resolved', 'column=' + dateResult.columnName + ' row=' + dateResult.rowNumber + ' rawValue=' + dateResult.value + ' transactionDate=' + transactionDate);

    return {
        value: transactionDate,
        rawValue: dateResult.value,
        rowNumber: dateResult.rowNumber,
        columnName: dateResult.columnName
    };
}

function getCurrentNetSuiteDate() {
    return nlapiDateToString(new Date(), 'date');
}

function getNetSuiteDateTimeValue(dateValue) {
    return nlapiDateToString(dateValue, 'datetimetz');
}

function debugLog(title, details) {
    if (LCL_DEBUG_LOGS) {
        nlapiLogExecution('DEBUG', title, details);
    }
}

function getErrorDetails(error) {
    if (!error) {
        return '';
    }

    var details = [];
    if (error.name) {
        details.push('name=' + error.name);
    }
    if (error.code) {
        details.push('code=' + error.code);
    }
    if (error.message) {
        details.push('message=' + error.message);
    }
    if (error.stack) {
        details.push('stack=' + error.stack);
    }

    if (details.length) {
        return details.join(' | ');
    }

    return String(error);
}

// ------------------------------------------------------------------
// LCL TRANSACTION CREATION
// ------------------------------------------------------------------

function setBodyField(record, fieldId, value) {
    if (value !== null && value !== undefined && value !== '') {
        debugLog('Setting body field', fieldId + '=' + value);
        record.setFieldValue(fieldId, String(value));
    }
}

function addLclItemLine(record, config, amount, documentNumber, fileName) {
    debugLog('Adding LCL item line', 'recordType=' + config.recordType + ' item=' + config.itemId + ' amount=' + amount + ' taxcode=' + LCL_TAX_CODE_ID + ' class=' + LCL_CLASS_ID + ' brand=' + LCL_BRAND_ID + ' documentNumber=' + documentNumber + ' fileName=' + fileName);
    record.selectNewLineItem('item');
    record.setCurrentLineItemValue('item', 'item', config.itemId);

    if (config.recordType === 'creditmemo') {
        record.setCurrentLineItemValue('item', 'price', '-1');
    }

    record.setCurrentLineItemValue('item', 'quantity', '1');
    record.setCurrentLineItemValue('item', 'units', LCL_UNITS_ID);
    record.setCurrentLineItemValue('item', 'rate', amount);
    record.setCurrentLineItemValue('item', 'amount', amount);
    record.setCurrentLineItemValue('item', 'taxcode', LCL_TAX_CODE_ID);
    record.setCurrentLineItemValue('item', 'class', LCL_CLASS_ID);
    record.setCurrentLineItemValue('item', 'cseg_mi_brand', LCL_BRAND_ID);

    if (config.setLineLocation) {
        record.setCurrentLineItemValue('item', 'location', LCL_LOCATION_ID);
    }

    if (config.setLineDescription) {
        record.setCurrentLineItemValue('item', 'description', documentNumber);
    }

    record.commitLineItem('item');
    debugLog('LCL item line committed', 'recordType=' + config.recordType + ' item=' + config.itemId + ' amount=' + amount);
}

function createCustomerRefundForCreditMemo(creditMemoId, customerId, documentNumber, transactionDate, amount) {
    if (!CUSTOMER_REFUND_ACCOUNT_ID) {
        throw nlapiCreateError('MISSING_REFUND_ACCOUNT', 'CUSTOMER_REFUND_ACCOUNT_ID is required to create Customer Refund', true);
    }

    var refundExternalId = buildEmailCaptureExternalId(getRefundCustomerKey(customerId) + '_refund', documentNumber);
var existingRefundId = findExistingEmailCaptureTransaction('customerrefund', refundExternalId);

if (existingRefundId) {
    nlapiLogExecution('AUDIT', 'Customer Refund already exists', 'externalId=' + refundExternalId + ' refundId=' + existingRefundId + ' documentNumber=' + documentNumber);
    return existingRefundId;
}

    debugLog('Creating Customer Refund', 'creditMemoId=' + creditMemoId + ' customerId=' + customerId + ' documentNumber=' + documentNumber + ' transactionDate=' + transactionDate + ' account=' + CUSTOMER_REFUND_ACCOUNT_ID + ' amount=' + amount);

    var refund = nlapiCreateRecord('customerrefund', { recordmode: 'dynamic' });
    var memo = 'EFT ' + documentNumber;
    setBodyField(refund, 'externalid', refundExternalId);
    setBodyField(refund, 'customer', customerId);
    setBodyField(refund, 'trandate', transactionDate);
    setBodyField(refund, 'account', CUSTOMER_REFUND_ACCOUNT_ID);
    setBodyField(refund, 'tranid', documentNumber);
    setBodyField(refund, 'memo', memo);
    setBodyField(refund, 'custbody_created_from_email_capture', 'T');
    setBodyField(refund, 'paymentmethod', 7); //EFT/ACH

    var line = refund.findLineItemValue('apply', 'doc', String(creditMemoId));
    debugLog('Customer Refund apply line lookup', 'creditMemoId=' + creditMemoId + ' line=' + line);

    if (line < 1) {
        throw nlapiCreateError('CREDIT_MEMO_NOT_ON_REFUND', 'Credit Memo ' + creditMemoId + ' was not found on Customer Refund apply sublist', true);
    }

    refund.selectLineItem('apply', line);
    refund.setCurrentLineItemValue('apply', 'apply', 'T');
    refund.setCurrentLineItemValue('apply', 'amount', amount);
    refund.commitLineItem('apply');

    debugLog('Submitting Customer Refund', 'creditMemoId=' + creditMemoId + ' documentNumber=' + documentNumber + ' amount=' + amount);
    var refundId = nlapiSubmitRecord(refund, true, true);
    nlapiLogExecution('AUDIT', 'Customer Refund created', 'refundId=' + refundId + ' creditMemoId=' + creditMemoId + ' documentNumber=' + documentNumber + ' amount=' + amount);

    return refundId;
}

function createLclTransaction(lclSubject, lclFile, fileName, transactionDate) {
    var config = getLclTransactionConfig(lclSubject.customerKey, lclSubject.transactionType);
    if (!config) {
        throw nlapiCreateError('LCL_CONFIG_MISSING', 'No LCL transaction config for type ' + lclSubject.transactionType, true);
    }

    var documentNumber = lclFile.documentNumber;


      var externalId = buildEmailCaptureExternalId(lclSubject.customerKey + '_' + lclSubject.transactionType, documentNumber);
var existingRecordId = findExistingEmailCaptureTransaction(config.recordType, externalId);

if (existingRecordId) {
    nlapiLogExecution('AUDIT', 'Email Capture transaction already exists', 'recordType=' + config.recordType + ' externalId=' + externalId + ' existingId=' + existingRecordId);

    var existingRefundId = null;
    if (config.recordType === 'creditmemo') {
        existingRefundId = createCustomerRefundForCreditMemo(existingRecordId, config.entityId, documentNumber, transactionDate, lclSubject.amount);
    }

    return {
        isLcl: true,
        created: true,
        id: existingRecordId,
        refundId: existingRefundId,
        recordType: config.recordType,
        label: config.label,
        tranid: documentNumber
    };
}

  
    debugLog('Creating LCL transaction', 'type=' + lclSubject.transactionType + ' recordType=' + config.recordType + ' documentNumber=' + documentNumber + ' fileName=' + fileName + ' amount=' + lclSubject.amount + ' timestamp=' + lclSubject.timestampText + ' transactionDate=' + transactionDate);
    var record = nlapiCreateRecord(config.recordType, { recordmode: 'dynamic' });
    var memo = 'EFT ' + documentNumber;
    setBodyField(record, 'externalid', externalId);
    setBodyField(record, 'entity', config.entityId);
    setBodyField(record, 'trandate', transactionDate);
    setBodyField(record, 'tranid', documentNumber);
    setBodyField(record, 'memo', memo);
    setBodyField(record, 'custbody_created_from_email_capture', 'T');
    setBodyField(record, 'currency', LCL_CURRENCY_ID);
    setBodyField(record, 'location', LCL_LOCATION_ID);
    setBodyField(record, 'account', config.accountId);
    setBodyField(record, 'custbody_report_timestamp', getNetSuiteDateTimeValue(lclSubject.timestampDate));
    setBodyField(record, config.referenceFieldId, documentNumber);

    addLclItemLine(record, config, lclSubject.amount, documentNumber, fileName);

    debugLog('Submitting LCL transaction', 'recordType=' + config.recordType + ' tranid=' + documentNumber + ' sourceFile=' + fileName + ' transactionDate=' + transactionDate);
    var recordId = nlapiSubmitRecord(record, true, true);
    nlapiLogExecution('AUDIT', 'LCL transaction created', config.label + ' id=' + recordId + ' tranid=' + documentNumber + ' sourceFile=' + fileName + ' amount=' + lclSubject.amount);
    var refundId = null;

    if (config.recordType === 'creditmemo') {
        refundId = createCustomerRefundForCreditMemo(recordId, config.entityId, documentNumber, transactionDate, lclSubject.amount);
    }

    return {
        isLcl: true,
        created: true,
        id: recordId,
        refundId: refundId,
        recordType: config.recordType,
        label: config.label,
        tranid: documentNumber
    };
}

function maybeCreateLclTransaction(subject, csvFile) {
    debugLog('Checking LCL transaction add-on', 'subject=' + subject + ' csvFile=' + csvFile.getName());
    var lclSubject = parseLclSubject(subject);
    if (!lclSubject) {
        debugLog('Not an LCL transaction subject', 'subject=' + subject);
        return { isLcl: false, created: false };
    }

    if (lclSubject.error) {
        nlapiLogExecution('ERROR', 'Invalid LCL transaction subject', lclSubject.error + ' subject=' + subject);
        return { isLcl: true, created: false };
    }

    if (lclSubject.skipTransaction && lclSubject.skipSecondBillCredit) {
        nlapiLogExecution('AUDIT', 'All transactions skipped', 'First and second amounts are zero or blank. subject=' + subject);
        return { isLcl: true, created: false, skipTransaction: true };
    }


    debugLog('LCL subject parsed', 'customer=' + lclSubject.customerName + ' transactionType=' + lclSubject.transactionType + ' timestamp=' + lclSubject.timestampText + ' amount=' + lclSubject.amount);
    var csvFileName = csvFile.getName();
    var lclFile = parseLclCsvFileName(csvFileName);
    if (lclFile.error) {
        nlapiLogExecution('ERROR', 'Invalid LCL CSV filename', lclFile.error + ' subject=' + subject);
        return { isLcl: true, created: false };
    }

    debugLog('LCL CSV filename parsed', 'transactionType=' + lclFile.transactionType + ' fileType=' + lclFile.fileType + ' documentNumber=' + lclFile.documentNumber + ' fileName=' + csvFileName);
    if (lclFile.transactionType !== lclSubject.transactionType) {
        nlapiLogExecution('ERROR', 'LCL subject/file type mismatch', 'subject=' + subject + ' fileName=' + csvFileName);
        return { isLcl: true, created: false };
    }

  if (lclFile.customerKey && lclFile.customerKey !== lclSubject.customerKey) {
    nlapiLogExecution('ERROR', 'Subject/file customer mismatch', 'subject=' + subject + ' fileName=' + csvFileName + ' subjectCustomer=' + lclSubject.customerName + ' fileCustomer=' + lclFile.customerName);
    return { isLcl: true, created: false };
}

    var config = getLclTransactionConfig(lclSubject.customerKey, lclSubject.transactionType);
    if (!config) {
        nlapiLogExecution('ERROR', 'Missing LCL transaction config', 'transactionType=' + lclSubject.transactionType + ' subject=' + subject);
        return { isLcl: true, created: false };
    }

    var transactionDateResult = getCsvTransactionDate(csvFile, config);
    if (transactionDateResult.error) {
        nlapiLogExecution('ERROR', 'CSV transaction date not resolved', transactionDateResult.error + ' subject=' + subject + ' fileName=' + csvFileName);
        return { isLcl: true, created: false };
    }

    var result = {
    isLcl: true,
    created: false,
    skipTransaction: false
};

if (lclSubject.skipTransaction) {
    nlapiLogExecution('AUDIT', 'Main transaction skipped', lclSubject.skipReason + ' subject=' + subject);
} else {
    result = createLclTransaction(lclSubject, lclFile, csvFileName, transactionDateResult.value);
}

if (lclSubject.skipSecondBillCredit === false) {
    var paybackResult = createPaybackVendorCredit(lclSubject, lclFile, csvFileName, transactionDateResult.value);
    result.isLcl = true;
    result.created = true;
    result.paybackVendorCreditId = paybackResult.id;
    result.paybackTranid = paybackResult.tranid;
} else {
    debugLog('Payback Vendor Credit skipped', lclSubject.secondSkipReason + ' subject=' + subject);
}

if (lclSubject.skipTransaction && lclSubject.skipSecondBillCredit) {
    result.skipTransaction = true;
}

return result;
}

function getMessageValue(message, methodName) {
    try {
        if (message && typeof message[methodName] === 'function') {
            return message[methodName]() || '';
        }
    } catch (e) {
        debugLog('Unable to read email sender method', methodName + ' error=' + getErrorDetails(e));
    }

    return '';
}

function getSenderEmail(message) {
    var rawSender = getMessageValue(message, 'getFrom') ||
            getMessageValue(message, 'getFromEmail') ||
            getMessageValue(message, 'getSender');
    var senderText = trim(rawSender);
    var emailMatch = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i.exec(senderText);

    return {
        raw: senderText,
        email: emailMatch ? emailMatch[0].toLowerCase() : senderText.toLowerCase()
    };
}

function isAllowedSender(message) {
    var sender = getSenderEmail(message);
    var emailParts = sender.email.split('@');
    var domain = emailParts.length === 2 ? emailParts[1] : '';
    var allowed = domain === ALLOWED_SENDER_DOMAIN;

    debugLog('Sender validation', 'rawSender=' + sender.raw + ' parsedEmail=' + sender.email + ' domain=' + domain + ' allowed=' + allowed);

    return {
        allowed: allowed,
        raw: sender.raw,
        email: sender.email,
        domain: domain
    };
}

// ------------------------------------------------------------------
// EMAIL CAPTURE ENTRY POINT
// ------------------------------------------------------------------

/**
 * Email Capture Plug-in entry point.
 * @param {nlobjEmail} message the received email
 * @param {nlobjRecord} newRecord stub record (e.g. Case) NetSuite would
 *                                otherwise create by default
 */
function process(message, newRecord) {
    try {
        var subject = message.getSubject();
        nlapiLogExecution('DEBUG', 'Email Capture triggered', 'subject=' + subject + ' targetFolderId=' + TARGET_FOLDER_ID);

      var senderValidation = isAllowedSender(message);
if (!senderValidation.allowed) {
    nlapiLogExecution('AUDIT', 'Email sender blocked', 'subject=' + subject + ' rawSender=' + senderValidation.raw + ' parsedEmail=' + senderValidation.email + ' parsedDomain=' + senderValidation.domain + ' allowedDomain=' + ALLOWED_SENDER_DOMAIN);
    return;
}

        var importConfig = resolveImportConfig(subject);
        if (!importConfig) {
            nlapiLogExecution('AUDIT', 'No matching import config for subject', subject);
            return;
        }
        debugLog('Import config resolved', 'keyword=' + importConfig.keyword + ' mappingId=' + importConfig.mappingId + ' description=' + importConfig.description);

        var csvFile = getCsvAttachment(message);
        if (!csvFile) {
            nlapiLogExecution('ERROR', 'Matched subject but no CSV attachment found', subject);
            return;
        }
        debugLog('CSV attachment ready', 'name=' + csvFile.getName());

        var fileId = saveAttachment(csvFile);
        debugLog('CSV attachment saved result', 'fileId=' + fileId + ' name=' + csvFile.getName());

        var savedCsvFile = nlapiLoadFile(fileId);
        debugLog('CSV file loaded for LCL parsing', 'fileId=' + fileId + ' name=' + savedCsvFile.getName());

        var lclResult = { isLcl: false, created: false };
        try {
            lclResult = maybeCreateLclTransaction(subject, savedCsvFile);
        } catch (lclError) {
            nlapiLogExecution('ERROR', 'LCL transaction creation failed', getErrorDetails(lclError));
            if (isLclTransactionSubject(subject)) {
                lclResult = { isLcl: true, created: false };
            }
        }

        if (lclResult && lclResult.isLcl && !lclResult.created && !lclResult.skipTransaction) {
    nlapiLogExecution('ERROR', 'CSV import not scheduled', 'LCL transaction was not created successfully. subject=' + subject + ' fileId=' + fileId + ' fileName=' + csvFile.getName());
    return;
}

        var importFileId = fileId;
        var importFileName = savedCsvFile.getName();
        var lclRemittanceImportResult = maybeCreateLclRemittanceUpdatedFile(subject, savedCsvFile);
        var metroCreditImportResult = { updated: false };

        if (lclRemittanceImportResult.error) {
            nlapiLogExecution('ERROR', 'CSV import not scheduled', 'LCL remittance combine failed. ' + lclRemittanceImportResult.error + ' subject=' + subject + ' fileId=' + fileId + ' fileName=' + csvFile.getName());
            return;
        }

        if (lclRemittanceImportResult.updated) {
            importFileId = lclRemittanceImportResult.fileId;
            importFileName = lclRemittanceImportResult.fileName;
        } else {
            metroCreditImportResult = maybeCreateMetroRemittanceUpdatedFile(subject, savedCsvFile);
        }

        if (metroCreditImportResult.error) {
            nlapiLogExecution('ERROR', 'CSV import not scheduled', 'Metro remittance credit filter failed. ' + metroCreditImportResult.error + ' subject=' + subject + ' fileId=' + fileId + ' fileName=' + csvFile.getName());
            return;
        }

        if (metroCreditImportResult.updated) {
            importFileId = metroCreditImportResult.fileId;
            importFileName = metroCreditImportResult.fileName;
        }

        triggerCsvImport(importFileId, importConfig.mappingId);

        if (newRecord) {
            var messageText = 'Auto-processed CSV import: ' + importConfig.description;
            if (lclResult && lclResult.created) {
    if (lclResult.id) {
        messageText += '; created ' + lclResult.recordType + ' ' + lclResult.id + ' tranid ' + lclResult.tranid;
    }
    if (lclResult.refundId) {
        messageText += '; created customerrefund ' + lclResult.refundId;
    }
    if (lclResult.paybackVendorCreditId) {
        messageText += '; created payback vendorcredit ' + lclResult.paybackVendorCreditId + ' tranid ' + lclResult.paybackTranid;
    }
}
            if (metroCreditImportResult.updated) {
                messageText += '; created filtered import file ' + importFileName + ' creditCount ' + metroCreditImportResult.creditCount;
            }
            if (lclRemittanceImportResult.updated) {
                messageText += '; created grouped import file ' + importFileName + ' combinedGroups ' + lclRemittanceImportResult.combinedGroupCount + ' removedNegativeRows ' + lclRemittanceImportResult.removedNegativeSourceRowCount;
            }
            newRecord.setFieldValue('incomingmessage', messageText);
        }

    } catch (e) {
        nlapiLogExecution('ERROR', 'Email Capture CSV import failed', getErrorDetails(e));
    }
}
