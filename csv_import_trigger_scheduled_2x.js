/**
 * CSV Import Trigger (Scheduled Script)
 *
 * Companion to email_capture_csv_import_1_0.js. The 1.0 Email Capture
 * Plug-in cannot call N/task directly, so it schedules this script via
 * nlapiScheduleScript, passing the saved file's internal id and the
 * target Saved CSV Import mapping id as script parameters. This script
 * does the actual import submission.
 *
 * SETUP REQUIRED:
 *  - Deploy as a Scheduled Script.
 *  - Create two script parameters on this deployment:
 *      custscript_import_file_id     (Free-Form Text or Integer)
 *      custscript_import_mapping_id  (Free-Form Text or Integer)
 *  - Make sure the script id / deployment id here match what's
 *    referenced as SCHEDULED_SCRIPT_ID / SCHEDULED_DEPLOYMENT_ID in
 *    email_capture_csv_import_1_0.js.
 *
 * @NApiVersion 2.1
 * @NScriptType ScheduledScript
 * @NModuleScope SameAccount
 */
define(['N/task', 'N/runtime', 'N/log','N/file'], function (task, runtime, log,file) {

    function execute(context) {
        var script = runtime.getCurrentScript();

        var fileId = script.getParameter({ name: 'custscript_import_file_id' });
        var mappingId = script.getParameter({ name: 'custscript_import_mapping_id' });

        if (!fileId || !mappingId) {
            log.error('Missing parameters', { fileId: fileId, mappingId: mappingId });
            return;
        }
            log.debug('parameters', { fileId: fileId, mappingId: mappingId });

        try {
           var csvData = file.load({
             id:fileId
           })
            var csvImportTask = task.create({
                taskType: task.TaskType.CSV_IMPORT,
                importFile: csvData,
                mappingId: parseInt(mappingId)
            });

            var taskId = csvImportTask.submit();
            log.audit('CSV import task submitted', {
                taskId: taskId,
                fileId: fileId,
                mappingId: mappingId
            });

        } catch (e) {
            log.error('CSV import task submission failed', e);
        }
    }

    return {
        execute: execute
    };

});
