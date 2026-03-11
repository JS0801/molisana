/**
* @NApiVersion 2.x
* @NScriptType ScheduledScript
*/
define(['N/task', 'N/file', 'N/log', 'N/runtime','N/search', 'N/record'],
function(task, file, log, runtime, search, record) {

  function execute(context) {
    try {

      var paramSearchId = runtime.getCurrentScript().getParameter({ name: 'custscript_mi_search' });
      var paramSearchId1 = runtime.getCurrentScript().getParameter({ name: 'custscript_mi_search1' });
      var paramSearchId2 = runtime.getCurrentScript().getParameter({ name: 'custscript_item_search' });

      var paramSearchId_avail = runtime.getCurrentScript().getParameter({ name: 'custscript_mi_avail_search' });
      var paramSearchId1_avail = runtime.getCurrentScript().getParameter({ name: 'custscript_inventory_avail_search' });
      var paramSearchId2_avail = runtime.getCurrentScript().getParameter({ name: 'custscript_item_list_avail' });
      var paramSearchId3_avail = runtime.getCurrentScript().getParameter({ name: 'custscript_item_last_billed_date' });


      var fileObj_avail = file.create({
        name: 'Avail_Tool_Assembly' + new Date() + '.csv',
        fileType: file.Type.CSV,
        encoding: file.Encoding.UTF8,
        folder: 423666,
        isOnline: true
      });

      var fileID_avail = fileObj_avail.save();
      log.debug('fileID', fileID_avail)

      var searchTask_avail = task.create({
        taskType: task.TaskType.SEARCH,
        savedSearchId: paramSearchId_avail,
        fileId: fileID_avail
      });

      var searchTaskId_avail = searchTask_avail.submit();
      log.audit("Search Export Task Submitted", "Task ID: " + searchTaskId_avail);

      var fileObj_avail1 = file.create({
        name: 'Avail_Tool_Inventory' + new Date() + '.csv',
        fileType: file.Type.CSV,
        encoding: file.Encoding.UTF8,
        folder: 423667,
        isOnline: true
      });

      var fileID_avail1 = fileObj_avail1.save();
      log.debug('fileID', fileID_avail1)

      var searchTask_avail1 = task.create({
        taskType: task.TaskType.SEARCH,
        savedSearchId: paramSearchId1_avail,
        fileId: fileID_avail1
      });

      var searchTaskId_avail1 = searchTask_avail1.submit();
      log.audit("Search Export Task Submitted", "Task ID: " + searchTaskId_avail1);

      var fileObj_avail2 = file.create({
        name: 'Avail_Tool_item_list' + new Date() + '.csv',
        fileType: file.Type.CSV,
        encoding: file.Encoding.UTF8,
        folder: 423668,
        isOnline: true
      });

      var fileID_avail2 = fileObj_avail2.save();
      log.debug('fileID', fileID_avail2)

      var searchTask_avail2 = task.create({
        taskType: task.TaskType.SEARCH,
        savedSearchId: paramSearchId2_avail,
        fileId: fileID_avail2
      });

      var searchTaskId_avail2 = searchTask_avail2.submit();
      log.audit("Search Export Task Submitted", "Task ID: " + searchTaskId_avail2);


      var fileObj_avail3 = file.create({
        name: 'Avail_Tool_BilledDate_' + new Date() + '.csv',
        fileType: file.Type.CSV,
        encoding: file.Encoding.UTF8,
        folder: 447164,
        isOnline: true
      });

      var fileID_avail3 = fileObj_avail3.save();
      log.debug('fileID', fileID_avail3)

      var searchTask_avail3 = task.create({
        taskType: task.TaskType.SEARCH,
        savedSearchId: paramSearchId3_avail,
        fileId: fileID_avail3
      });

      var searchTaskId_avail3 = searchTask_avail3.submit();
      log.audit("Search Export Task Submitted", "Task ID: " + searchTaskId_avail3);

      //-------------------------

      var fileObj = file.create({
        name: 'RR_Tool_Assembly' + new Date() + '.csv',
        fileType: file.Type.CSV,
        encoding: file.Encoding.UTF8,
        folder: 402335,
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

      var fileObj1 = file.create({
        name: 'RR_Tool_Inventory' + new Date() + '.csv',
        fileType: file.Type.CSV,
        encoding: file.Encoding.UTF8,
        folder: 402334,
        isOnline: true
      });

      var fileID1 = fileObj1.save();
      log.debug('fileID1', fileID1)

      var searchTask1 = task.create({
        taskType: task.TaskType.SEARCH,
        savedSearchId: paramSearchId1,
        fileId: fileID1
      });

      var searchTaskId1 = searchTask1.submit();
      log.audit("Search Export Task Submitted", "Task ID1: " + searchTaskId1);

      var fileObj2 = file.create({
        name: 'RR_Tool_Item' + new Date() + '.csv',
        fileType: file.Type.CSV,
        encoding: file.Encoding.UTF8,
        folder: 413248,
        isOnline: true
      });

      var fileID2 = fileObj2.save();
      log.debug('fileID', fileID2)

      var searchTask2 = task.create({
        taskType: task.TaskType.SEARCH,
        savedSearchId: paramSearchId2,
        fileId: fileID2
      });

      var searchTaskId2 = searchTask2.submit();
      log.audit("Search Export Task Submitted", "Task ID: " + searchTaskId2);




    } catch (e) {
      log.error("Error Executing Scheduled Script", e.message);
    }
  }

  return {
    execute: execute
  };
});
