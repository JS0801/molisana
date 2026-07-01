function afterSubmitFunction() {
  // Get the internal ID of the current record
  var currentRecordId = nlapiGetRecordId();
  nlapiLogExecution('DEBUG', 'Current Record ID', 'Current Record ID: ' + currentRecordId);
  
  // Load the inbound shipment record using the current record's internal ID
  var inboundShipment = nlapiLoadRecord('inboundshipment', currentRecordId);
  nlapiLogExecution('DEBUG', 'inboundShipment Record Loaded', 'Record loaded successfully');
  
  // Retrieve the value of custrecord_mi_ds_script_obj
  var scriptObjValue = inboundShipment.getFieldValue('custrecord_mi_ds_script_obj');
  
  if (scriptObjValue) {
    // Parse the JSON string into an array of objects
    var itemsArray = JSON.parse(scriptObjValue);
    
    for (var key in itemsArray) {
      if (itemsArray.hasOwnProperty(key)) {
        var costCode = key;
        nlapiLogExecution('DEBUG', 'costCode', 'Cost Code: ' + costCode);

        var line_code = '';
        var SEPARATOR = '\u0005';
        var codeLine = [];
        
        // Loop through each item in the array
        for (var i = 0; i < itemsArray[costCode].length; i++) {
          var itemObj = itemsArray[costCode][i];
          
          // Use findLineItemValue to find the line number where itemID matches
          var lineNumber = inboundShipment.findLineItemValue('items', 'itemid', itemObj.itemID);
          
          if (lineNumber > 0) { // lineNumber is 1-based index, 0 means not found
            nlapiLogExecution('DEBUG', 'Item Found', 'Item ID: ' + itemObj.itemID + ' found at line: ' + lineNumber);
            
            // Get the internal ID of the line (if needed)
            var lineUniqueKey = inboundShipment.getLineItemValue('items', 'id', lineNumber);
            nlapiLogExecution('DEBUG', 'Line Unique Key', 'Line Unique Key for item ID ' + itemObj.itemID + ': ' + lineUniqueKey);

            if (lineUniqueKey) {
              if (line_code.indexOf(lineUniqueKey) == -1) {
                line_code += (i != itemsArray[costCode].length - 1) ? lineUniqueKey + SEPARATOR : lineUniqueKey;
                codeLine.push(lineUniqueKey);
              }            
            }
          } else {
            nlapiLogExecution('DEBUG', 'Item Not Found', 'Item ID: ' + itemObj.itemID + ' was not found in the sublist.');
          }
        }

        nlapiLogExecution('DEBUG', 'line_code Final', 'Final line_code: ' + line_code);

        var costCodeID = costCode.split("-")[0];
        var amount = costCode.split("-")[1];
        var currency = costCode.split("-")[2];
        var exchangeRate = nlapiExchangeRate(currency, 1);
        

        // Ensure that the sublist is cleared before adding new lines
        inboundShipment.selectNewLineItem('landedcost');
        inboundShipment.setCurrentLineItemValue('landedcost', 'landedcostshipmentitems', codeLine);
        inboundShipment.setCurrentLineItemValue('landedcost', 'landedcostallocationmethod', 'QUANTITY');
        inboundShipment.setCurrentLineItemValue('landedcost', 'landedcostcostcategory', costCodeID);
        inboundShipment.setCurrentLineItemValue('landedcost', 'landedcostamount', amount);
        inboundShipment.setCurrentLineItemValue('landedcost', 'landedcostcurrency', currency);
        inboundShipment.setCurrentLineItemValue('landedcost', 'landedcostexchangerate', exchangeRate);
        inboundShipment.setCurrentLineItemValue('landedcost', 'landedcosteffectivedate',new Date());
        inboundShipment.commitLineItem('landedcost');        
      }
    }
  }
  
  // Save the record
  nlapiLogExecution('DEBUG', 'inboundShipment Record Updated', 'Record updated successfully');

  // Clear the script object field to prevent reprocessing
  inboundShipment.setFieldValue('custrecord_mi_ds_script_obj', '');
  nlapiSubmitRecord(inboundShipment);
}
