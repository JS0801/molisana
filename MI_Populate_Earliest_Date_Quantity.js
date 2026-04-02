/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 */
define(['N/search','N/record','N/format','N/log'], function (search, record, format, log) {

  var ITEM_RECORD_TYPE = 'lotnumberedassemblyitem';
  var FLD_PORT_ETA = 'custitem_mi_earliest_port_date';
  var FLD_EXP_QTY  = 'custitem_mi_earliest_expected_quantity';

  function getInputData() {
    // Items that appear on an inbound shipment with status = In Transit
    var s = search.create({
      type: 'inboundshipment',
      filters: [['status','anyof','inTransit']],
      columns: [ search.createColumn({ name: 'item', summary: 'GROUP' }) ]
    });

    var items = [];
    s.run().each(function (res) {
      var itemId = res.getValue({ name: 'item', summary: 'GROUP' });
      if (itemId) items.push({ itemId: String(itemId) });
      return true;
    });
    return items;
  }

  function map(context) {
    var row = JSON.parse(context.value);
    var itemId = Number(row.itemId);
    if (!itemId) return;

    var earliest = findEarliestInboundForItem(itemId);
    if (!earliest) return;

    // Load the item and read current values
    var itemRec;
    try {
      itemRec = record.load({ type: ITEM_RECORD_TYPE, id: itemId, isDynamic: false });
    } catch (e) {
      log.error('Item load failed', { itemId: itemId, error: e });
      return;
    }

    var oldDateVal = itemRec.getValue({ fieldId: FLD_PORT_ETA });
    var oldQtyRaw  = itemRec.getValue({ fieldId: FLD_EXP_QTY });
    var oldQty     = normalizeNumber(oldQtyRaw);

    var oldDateStr = null;
    try {
      if (oldDateVal) {
        if (Object.prototype.toString.call(oldDateVal) === '[object Date]') {
          oldDateStr = format.format({ value: oldDateVal, type: format.Type.DATE });
        } else {
          var tmp = format.parse({ value: String(oldDateVal), type: format.Type.DATE });
          oldDateStr = format.format({ value: tmp, type: format.Type.DATE });
        }
      }
    } catch (e) {}

    var dateChanged = (oldDateStr !== earliest.etaStr);
    var qtyChanged  = !numbersEqual(oldQty, earliest.expectedQty);
    if (!dateChanged && !qtyChanged) return;

    // Write changes
    try {
      record.submitFields({
        type: ITEM_RECORD_TYPE,
        id: itemId,
        values: (function(){ var o={}; o[FLD_PORT_ETA]=earliest.etaObj; o[FLD_EXP_QTY]=earliest.expectedQty; return o; })(),
        options: { enableSourcing: false, ignoreMandatoryFields: true }
      });

      // Minimal log showing what changed
      log.audit('Item updated', {
        itemId: itemId,
        inboundId: earliest.inboundId,
        old_port_date: oldDateStr,
        new_port_date: earliest.etaStr,
        old_expected_qty: oldQty,
        new_expected_qty: earliest.expectedQty
      });
    } catch (e) {
      log.error('submitFields failed', { itemId: itemId, error: e });
    }
  }

  function summarize(summary) {
    if (summary.inputSummary.error) log.error('Input Error', summary.inputSummary.error);
    summary.mapSummary.errors.iterator().each(function (key, e) {
      log.error('Map Error ' + key, e);
      return true;
    });
  }

  // --- helpers ---

  function findEarliestInboundForItem(itemId) {
    var rs = search.create({
      type: 'inboundshipment',
      filters: [
        ['status','anyof','inTransit'],'AND',
        ['item','anyof', itemId]
      ],
      columns: [
        search.createColumn({ name: 'custrecord_port_eta', sort: search.Sort.ASC }),
        search.createColumn({ name: 'quantityexpected' }),
        search.createColumn({ name: 'internalid' })
      ]
    }).run().getRange({ start: 0, end: 1 });

    if (!rs || !rs.length) return null;

    var etaRaw = rs[0].getValue({ name: 'custrecord_port_eta' });
    if (!etaRaw) return null;

    var etaObj, etaStr;
    try {
      etaObj = format.parse({ value: etaRaw, type: format.Type.DATE });
      etaStr = format.format({ value: etaObj, type: format.Type.DATE });
    } catch (e) {
      log.error('ETA parse/format failed', { itemId: itemId, raw: etaRaw, error: e });
      return null;
    }

    var qty = normalizeNumber(rs[0].getValue({ name: 'quantityexpected' }));
    var inboundId = rs[0].getValue({ name: 'internalid' });

    return { etaObj: etaObj, etaStr: etaStr, expectedQty: qty, inboundId: inboundId };
  }

  function normalizeNumber(v) {
    if (v === null || typeof v === 'undefined' || v === '') return null;
    var n = Number(v);
    return isNaN(n) ? null : n;
  }

  function numbersEqual(a, b) {
    if (a === null && b === null) return true;
    if (a === null || b === null) return false;
    return Number(a) === Number(b);
  }

  return { getInputData: getInputData, map: map, summarize: summarize };
});
