/**
 * @NApiVersion 2.x
 * @NScriptType UserEventScript
 */
define(['N/record', 'N/log'], function(record, log) {

    function afterSubmit(context) {
        if (context.type !== context.UserEventType.CREATE) return;

      try {

        var newRecord = context.newRecord;

        // Get the header field value
        var headerInboundShipment = newRecord.getValue({ fieldId: 'custbody_vendbill_related_to_inbship' });
        var venRefNum = newRecord.getValue({fieldId: 'tranid'})

        if (!headerInboundShipment) {
            log.debug('No Header Value', 'No value found for custbody_vendbill_related_to_inbship');
            return;
        }

        // Load the record in edit mode
        var vendorBill = record.load({
            type: 'vendorbill',
            id: newRecord.id,
            isDynamic: false
        });

        // Get the number of lines in the item sublist
        var lineCount = vendorBill.getLineCount({ sublistId: 'item' });

        // Loop through the lines in reverse to safely remove lines
        for (var i = lineCount - 1; i >= 0; i--) {
            var lineInboundShipment = vendorBill.getSublistValue({
                sublistId: 'item',
                fieldId: 'custcol_mi_related_inbound',
                line: i
            });

            var venRef = vendorBill.getSublistValue({
                sublistId: 'item',
                fieldId: 'custcol_mi_vendor_ref_number',
                line: i
            });

            // If line's field does not match the header field, remove the line
            if (lineInboundShipment != headerInboundShipment || venRef != venRefNum) {
                vendorBill.removeLine({
                    sublistId: 'item',
                    line: i,
                    ignoreRecalc: true
                });
                log.debug('Line Removed', 'Line ' + i + ' removed due to mismatch.');
            }
        }

        // Save the updated record
        vendorBill.save({ ignoreMandatoryFields: true });

        log.debug('Vendor Bill Updated', 'Non-matching lines removed successfully.');
        
        
      } catch (error) {
        log.error('error',error)
      }
    }

    return {
        afterSubmit: afterSubmit
    };
});
