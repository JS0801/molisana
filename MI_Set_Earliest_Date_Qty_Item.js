/**
 * @NApiVersion 2.x
 * @NScriptType UserEventScript
 */

define(['N/record', 'N/search', 'N/log', 'N/format'], function (record, search, log, format) {
  function afterSubmit(context) {
    if (context.type !== context.UserEventType.CREATE && context.type !== context.UserEventType.EDIT) return;
    //DS GIT Test
    try {
      var shipment = context.newRecord;
      var itemLineCount = shipment.getLineCount({ sublistId: 'items' });
      log.debug('Item Line Count', itemLineCount);

      for (var i = 0; i < itemLineCount; i++) {
        var itemId = shipment.getSublistValue({
          sublistId: 'items',
          fieldId: 'itemid',
          line: i
        });

        log.debug('Processing Item ID', itemId);
        if (!itemId) continue;

        // Load saved search fresh each iteration
        var shipmentSearch = search.load({ id: 'customsearch3533' });

        // Remove any existing item filter if needed
        shipmentSearch.filters = shipmentSearch.filters.filter(function (f) {
          return f.name !== 'item';
        });

        // Add dynamic item filter
        shipmentSearch.filters.push(search.createFilter({
          name: 'item',
          operator: search.Operator.ANYOF,
          values: [itemId]
        }));

        var results = shipmentSearch.run().getRange({ start: 0, end: 1 });

        if (results.length > 0) {
          var minDate = results[0].getValue({ name: 'custrecord_port_eta', summary: "MIN" });
          var expectedQty = results[0].getValue({ name: 'quantityexpected', summary: "MAX" });

          log.debug('Found Data', { itemId: itemId, minDate: minDate, expectedQty: expectedQty });

          if (minDate) {
            var formattedDate = format.parse({ value: minDate, type: format.Type.DATE });

            var itemRec = record.load({
              type: "lotnumberedassemblyitem",
              id: itemId,
              isDynamic: false
            });

            itemRec.setValue({
              fieldId: 'custitem_mi_earliest_port_date',
              value: formattedDate
            });

            itemRec.setValue({
              fieldId: 'custitem_mi_earliest_expected_quantity',
              value: expectedQty
            });

            var savedId = itemRec.save();
            log.audit('Item Updated', 'Item ID ' + savedId + ' - Port ETA: ' + minDate + ', Quantity: ' + expectedQty);
          }
        } else {
          log.debug('No Results for Item', itemId);
        }
      }

    } catch (e) {
      log.error('Error in afterSubmit', e);
    }
  }

  return {
    afterSubmit: afterSubmit
  };
});
