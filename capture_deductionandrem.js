/**
 * Email Capture Plug-in (SuiteScript 1.0)
 *
 * Existing behavior:
 *   1. Matches the subject line against SUBJECT_IMPORT_MAP
 *   2. Saves the CSV attachment to the File Cabinet
 *   3. Schedules the 2.x CSV import trigger script
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
 */

// ------------------------------------------------------------------
// CONFIG
// ------------------------------------------------------------------
var SUBJECT_IMPORT_MAP = [
    { keyword: 'Deduction',  mappingId: '157', description: 'Deduction CSV import' },
    { keyword: 'Remittance', mappingId: '156', description: 'Remittance CSV import' }
];

var TARGET_FOLDER_ID = 427297;

var SCHEDULED_SCRIPT_ID = 'customscript_mi_import_bills_and_payment';
var SCHEDULED_DEPLOYMENT_ID = 'customdeploy_mi_import_bills_and_payment';

var LCL_DEDUCTION_TYPE = 'deduction';
var LCL_REMITTANCE_TYPE = 'remittance';

var LCL_TRANSACTION_CONFIG = {};

LCL_TRANSACTION_CONFIG[LCL_DEDUCTION_TYPE] = {
    label: 'LCL deduction vendor credit',
    recordType: 'vendorcredit',
    entityId: '11437',
    accountId: '2059',
    itemId: '6733',
    fileType: 'Deduction',
    referenceFieldId: 'custbody_note_to_vendor',
    setLineLocation: true,
    setLineDescription: false
};

LCL_TRANSACTION_CONFIG[LCL_REMITTANCE_TYPE] = {
    label: 'LCL remittance credit memo',
    recordType: 'creditmemo',
    entityId: '30',
    accountId: '119',
    itemId: '6808',
    fileType: 'Remittance',
    referenceFieldId: 'custbody_2663_reference_num',
    setLineLocation: false,
    setLineDescription: true
};

var LCL_SUBSIDIARY_ID = '2';
var LCL_CURRENCY_ID = '1';
var LCL_LOCATION_ID = '315';
var LCL_CLASS_ID = '319';
var LCL_BRAND_ID = '227';
var LCL_TAX_CODE_ID = '16';
var LCL_UNITS_ID = '35';

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

    for (var i = 0; i < attachments.length; i++) {
        var name = attachments[i].getName() || '';
        if (name.toLowerCase().indexOf('.csv') !== -1) {
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
    if (!/^\s*LCL\s+(Deductions|Remittances)_/i.test(subject || '')) {
        return null;
    }

    var parts = String(subject).split('_');
    if (parts.length !== 3) {
        return { error: 'Expected exactly 3 underscore-delimited parts: type_timestamp_amount' };
    }

    var typeText = trim(parts[0]);
    var timestampText = trim(parts[1]);
    var amountText = trim(parts[2]);
    var transactionType = null;

    if (/^LCL\s+Deductions$/i.test(typeText)) {
        transactionType = LCL_DEDUCTION_TYPE;
    } else if (/^LCL\s+Remittances$/i.test(typeText)) {
        transactionType = LCL_REMITTANCE_TYPE;
    } else {
        return { error: 'Unsupported LCL subject type: ' + typeText };
    }

    var timestampDate = parseUtcTimestamp(timestampText);
    if (!timestampDate) {
        return { error: 'Invalid timestamp. Expected YYYY-MM-DD HH:MM:SSZ, received: ' + timestampText };
    }

    var amount = parseAmount(amountText);
    if (!amount) {
        return { error: 'Invalid amount. Expected a positive decimal amount, received: ' + amountText };
    }

    return {
        transactionType: transactionType,
        timestampText: timestampText,
        timestampDate: timestampDate,
        amount: amount
    };
}

function parseLclCsvFileName(fileName) {
    var match = /^(Deduction|Remittance)_([0-9]+)\.csv$/i.exec(trim(fileName));
    if (!match) {
        return { error: 'Invalid LCL CSV filename. Expected Deduction_number.csv or Remittance_number.csv, received: ' + fileName };
    }

    return {
        transactionType: /^Deduction$/i.test(match[1]) ? LCL_DEDUCTION_TYPE : LCL_REMITTANCE_TYPE,
        fileType: match[1],
        documentNumber: match[2]
    };
}

function getCurrentNetSuiteDate() {
    return nlapiDateToString(new Date(), 'date');
}

function getNetSuiteDateTimeValue(dateValue) {
    return nlapiDateToString(dateValue, 'datetimetz');
}

// ------------------------------------------------------------------
// LCL TRANSACTION CREATION
// ------------------------------------------------------------------

function setBodyField(record, fieldId, value) {
    if (value !== null && value !== undefined && value !== '') {
        record.setFieldValue(fieldId, String(value));
    }
}

function addLclItemLine(record, config, amount, fileName) {
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
        record.setCurrentLineItemValue('item', 'description', fileName);
    }

    record.commitLineItem('item');
}

function createLclTransaction(lclSubject, fileName) {
    var config = LCL_TRANSACTION_CONFIG[lclSubject.transactionType];
    if (!config) {
        throw nlapiCreateError('LCL_CONFIG_MISSING', 'No LCL transaction config for type ' + lclSubject.transactionType, true);
    }

    var record = nlapiCreateRecord(config.recordType, { recordmode: 'dynamic' });
    var memo = 'Auto-created from LCL email attachment ' + fileName;

  //  setBodyField(record, 'subsidiary', LCL_SUBSIDIARY_ID);
    setBodyField(record, 'entity', config.entityId);
    setBodyField(record, 'trandate', getCurrentNetSuiteDate());
    setBodyField(record, 'tranid', fileName);
    setBodyField(record, 'memo', memo);
    setBodyField(record, 'currency', LCL_CURRENCY_ID);
    setBodyField(record, 'location', LCL_LOCATION_ID);
    setBodyField(record, 'account', config.accountId);
    setBodyField(record, 'custbody_report_timestamp', getNetSuiteDateTimeValue(lclSubject.timestampDate));
    setBodyField(record, config.referenceFieldId, fileName);

    addLclItemLine(record, config, lclSubject.amount, fileName);

    var recordId = nlapiSubmitRecord(record, true, true);
    nlapiLogExecution('AUDIT', 'LCL transaction created', config.label + ' id=' + recordId + ' tranid=' + fileName + ' amount=' + lclSubject.amount);

    return {
        id: recordId,
        recordType: config.recordType,
        label: config.label
    };
}

function maybeCreateLclTransaction(subject, csvFile) {
    var lclSubject = parseLclSubject(subject);
    if (!lclSubject) {
        return null;
    }

    if (lclSubject.error) {
        nlapiLogExecution('ERROR', 'Invalid LCL transaction subject', lclSubject.error + ' subject=' + subject);
        return null;
    }

    var csvFileName = csvFile.getName();
    var lclFile = parseLclCsvFileName(csvFileName);
    if (lclFile.error) {
        nlapiLogExecution('ERROR', 'Invalid LCL CSV filename', lclFile.error + ' subject=' + subject);
        return null;
    }

    if (lclFile.transactionType !== lclSubject.transactionType) {
        nlapiLogExecution('ERROR', 'LCL subject/file type mismatch', 'subject=' + subject + ' fileName=' + csvFileName);
        return null;
    }

    return createLclTransaction(lclSubject, csvFileName);
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
        nlapiLogExecution('DEBUG', 'Email Capture triggered', 'subject=' + subject);

        var importConfig = resolveImportConfig(subject);
        if (!importConfig) {
            nlapiLogExecution('AUDIT', 'No matching import config for subject', subject);
            return;
        }

        var csvFile = getCsvAttachment(message);
        if (!csvFile) {
            nlapiLogExecution('ERROR', 'Matched subject but no CSV attachment found', subject);
            return;
        }

        var fileId = saveAttachment(csvFile);
        triggerCsvImport(fileId, importConfig.mappingId);

        var lclResult = null;
        try {
            lclResult = maybeCreateLclTransaction(subject, csvFile);
        } catch (lclError) {
            nlapiLogExecution('ERROR', 'LCL transaction creation failed', lclError.toString());
        }

        if (newRecord) {
            var messageText = 'Auto-processed CSV import: ' + importConfig.description;
            if (lclResult) {
                messageText += '; created ' + lclResult.recordType + ' ' + lclResult.id;
            }
            newRecord.setFieldValue('incomingmessage', messageText);
        }

    } catch (e) {
        nlapiLogExecution('ERROR', 'Email Capture CSV import failed', e.toString());
    }
}
