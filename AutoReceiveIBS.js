/*
**********************************************************************
*
*
* Author:      DHRUV SONI
*
***********************************************************************
* Script Description:
* This is User Event after submit and deployed on Sales order. context = Suitelet Only and ON CREATE.
* Script will verify on hand quantity based on the locations and assign peorper location to each items.
***********************************************************************
*/
/**
*@NApiVersion 2.x
*@NScriptType UserEventScript
*@NModuleScope SameAccount
*/
define(['N/error','N/log','N/record','N/search','N/format'],
function(error,log,record,search,format) {

  function afterSubmit(context) {

    var newRecord = context.newRecord;      
    var recID = newRecord.id;

    var validation = newRecord.getValue('custrecord_ds_receive_shipment')
    var recLocation = newRecord.getValue('custrecord_ds_receiving_location')

    var status = newRecord.getValue('shipmentstatus')
    log.debug('status', status)
    log.debug('validation', validation)

    if (!validation || status == 'received' || !recLocation) {
      return;
    }


    if(context.type == context.UserEventType.EDIT){

      try {
        var InboundReceiveObj = record.load({
          type: 'receiveinboundshipment',
          id: recID,
          isDynamic: false
        });

        InboundReceiveObj.setValue("shipmentstatus", "received");

        var InbReceiveobj_item_count = InboundReceiveObj.getLineCount({ sublistId: 'receiveitems' });

        log.debug("InbReceiveobj_item_count", InbReceiveobj_item_count);

        for (var i=0;i<InbReceiveobj_item_count;i++) {

          InboundReceiveObj.setSublistValue({ "sublistId": "receiveitems", "fieldId": "receiveitem", line:i, "value": true });
          InboundReceiveObj.setSublistValue({ "sublistId": "receiveitems", "fieldId": "receivinglocation", line:i, "value": recLocation });


          var quantityremaining = InboundReceiveObj.getSublistValue({ "sublistId": "receiveitems", "fieldId": "quantityremaining",line:i});
          log.debug("quantityremaining", quantityremaining);

          var itemID = InboundReceiveObj.getSublistValue({"sublistId": "receiveitems", "fieldId": "item",line:i});
          log.debug("itemID", itemID);
          
          InboundReceiveObj.setSublistValue({ "sublistId": "receiveitems", "fieldId": "quantitytobereceived",line:i, "value": quantityremaining });

          var subrecord = InboundReceiveObj.getSublistSubrecord({
            sublistId: 'receiveitems',
            fieldId: 'inventorydetail',
            line:i,
            ignoreRecalc: true
          });

          subrecord.setSublistValue({
            sublistId: 'inventoryassignment',
            fieldId: 'receiptinventorynumber',
            line:0,
            value: itemID//issueinventorynumber
          });

          subrecord.setSublistValue({
            sublistId: 'inventoryassignment',
            fieldId: 'inventorystatus',
            line:0,
            value: 1
          });

          subrecord.setSublistValue({
            sublistId: 'inventoryassignment',
            fieldId: 'quantity',
            line:0,
            value: quantityremaining
          });
        }

    var IRIDs =  InboundReceiveObj.save();
    log.debug( { 'IRIDs': 'IRIDs', 'IRIDs': IRIDs } );


      } catch( e ) {
        log.debug( { 'title': 'error', 'details': e } );
        return { 'error': { 'type': e.type, 'name': e.name, 'message': e.message } }
      }
    }

  }

  function isEmpty(value) {
    if (value === null) {
      return true;
    } else if (value === undefined) {
      return true;
    } else if (value === '') {
      return true;
    } else if (value === ' ') {
      return true;
    } else if (value === 'null') {
      return true;
    } else {
      return false;
    }
  }

  return {
    afterSubmit: afterSubmit
  };
});
