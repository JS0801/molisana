/**
* @NApiVersion 2.x
* @NScriptType ScheduledScript
*/
define(['N/task', 'N/file', 'N/log', 'N/runtime'],
function(task, file, log, runtime) {
  
  function execute(context) {
    try {
      
      var paramSearchId = runtime.getCurrentScript().getParameter({ name: 'custscript_bi_search' });
      var fileName = runtime.getCurrentScript().getParameter({ name: 'custscript_file_name' });
      var folderID = runtime.getCurrentScript().getParameter({ name: 'custscript_file_folder_id' });

      log.debug('Param', 'paramSearchId: '+ paramSearchId + ' -- fileName: ' + fileName + ' -- folderID: ' + folderID)
      
      var fileObj = file.create({
        name: fileName + '.csv',
        fileType: file.Type.CSV,
        encoding: file.Encoding.UTF8,
        folder: folderID,
        isOnline: true
      });

      var fileID = fileObj.save();
      log.debug('fileID', fileID)

      var searchTask = task.create({
        taskType: task.TaskType.SEARCH,
        savedSearchId: paramSearchId,
        fileId: fileID
      });

      var searchTaskId = searchTask.submit();
      log.audit("Search Export Task Submitted", "Task ID: " + searchTaskId);
      
    } catch (e) {
      log.error("Error Executing Scheduled Script", e.message);
    }
  }
  
  return {
    execute: execute
  };
});
