/**
 * Email Capture Plug-in (SuiteScript 1.0)
 *
 * The "Email Capture Plug-in" interface only runs as SuiteScript 1.0 -
 * NetSuite has not exposed a 2.x version of this interface. Because 1.0
 * has no N/task module, this script cannot submit the CSV Import task
 * itself. Instead it:
 *   1. Matches the subject line against the keyword table below
 *   2. Grabs the CSV attachment and saves it to the File Cabinet,
 *      keeping the original filename exactly as received
 *   3. Hands off to a 2.x Scheduled Script (see the companion file
 *      csv_import_trigger_scheduled_2x.js) via nlapiScheduleScript,
 *      passing the file id and mapping id as script parameters -
 *      that scheduled script is what actually calls N/task.
 *
 * SETUP REQUIRED:
 *  - Deploy this as a Plug-in Implementation script, Interface =
 *    "Email Capture Plug-in".
 *  - Deploy the companion 2.x scheduled script separately and update
 *    SCHEDULED_SCRIPT_ID / SCHEDULED_DEPLOYMENT_ID below to match.
 *  - This interface only fires within the CRM email-to-case pipeline
 *    (Setup > Support > Support Email Issue / Case Rules) - confirm
 *    that's the mailbox these CSVs are actually arriving at.
 *
 * VERIFY BEFORE DEPLOY:
 *  - Exact function signature NetSuite expects for this interface in
 *    your account version (I've used the commonly documented
 *    process(message, newRecord) shape below, but confirm against the
 *    Interface's method signature shown when you create the Plug-in
 *    Implementation script record).
 *  - nlobjEmail attachment API (getAttachments, getName) matches what's
 *    available in your account.
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

        if (newRecord) {
            newRecord.setFieldValue('incomingmessage', 'Auto-processed CSV import: ' + importConfig.description);
        }

    } catch (e) {
        nlapiLogExecution('ERROR', 'Email Capture CSV import failed', e.toString());
    }
}
