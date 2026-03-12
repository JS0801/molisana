/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 */
define(['N/search', 'N/record', 'N/runtime', 'N/log'], (search, record, runtime, log) => {

  const YEAR_FIELDS = [
    { year: '2021', fieldId: 'custitem_average_cost_2021' },
    { year: '2022', fieldId: 'custitem_average_cost_2022' },
    { year: '2023', fieldId: 'custitem_average_cost_2023' },
    { year: '2024', fieldId: 'custitem_average_cost_2024' },
    { year: '2025', fieldId: 'custitem_average_cost_2025' }
  ];
  const EPSILON = 0.0001;

  function toNumber(v) {
    if (v === null || v === '' || typeof v === 'undefined') return null;
    var n = Number(v);
    // keep it numeric (rounded to 2dp) so math works in approxEqual
    return isFinite(n) ? Math.round(n * 100) / 100 : null;
  }

  function approxEqual(a, b, eps = EPSILON) {
    if (a === null || b === null) return false;
    const na = Number(a), nb = Number(b);
    if (!isFinite(na) || !isFinite(nb)) return false;
    return Math.abs(na - nb) <= eps;
  }

  // UPDATED: handle lot-numbered types
  function nsItemTypeFromText(itemTypeText, isLot) {
    if (!itemTypeText) return null;
    const t = String(itemTypeText).toLowerCase();

    const isInventory = (t.indexOf('invtpart') > -1) || t === 'inventory item' || t === 'inventoryitem';
    const isAssembly  = (t.indexOf('assembly') > -1);

    if (isLot === true || String(isLot).toUpperCase() === 'T') {
      if (isInventory) return record.Type.LOT_NUMBERED_INVENTORY_ITEM;
      if (isAssembly)  return record.Type.LOT_NUMBERED_ASSEMBLY_ITEM;
    } else {
      if (isInventory) return record.Type.INVENTORY_ITEM;
      if (isAssembly)  return record.Type.ASSEMBLY_ITEM;
    }
    return null;
  }

  // ---- getInputData ----
  const getInputData = () => {
    log.audit('getInputData', 'Starting vendor bill search');

    const s = search.create({
      type: 'vendorbill',
      filters: [
        ['type', 'anyof', 'VendBill'], 'AND',
        ['mainline', 'is', 'F'], 'AND',
        ['taxline', 'is', 'F'], 'AND',
        ['shipping', 'is', 'F'], 'AND',
        ['cogs', 'is', 'F'], 'AND',
        ['item.type', 'anyof', 'InvtPart', 'Assembly']
      ],
      columns: [
        search.createColumn({ name: 'internalid', join: 'item', summary: 'GROUP' }),
        search.createColumn({ name: 'type', join: 'item', summary: 'GROUP' }),
        search.createColumn({ name: 'custitem_average_cost_2021', join: 'item', summary: 'MAX' }),
        search.createColumn({ name: 'custitem_average_cost_2022', join: 'item', summary: 'MAX' }),
        search.createColumn({ name: 'custitem_average_cost_2023', join: 'item', summary: 'MAX' }),
        search.createColumn({ name: 'custitem_average_cost_2024', join: 'item', summary: 'MAX' }),
        search.createColumn({ name: 'custitem_average_cost_2025', join: 'item', summary: 'MAX' }),
        search.createColumn({ name: 'custitem_average_cost_last_52_weeks', join: 'item', summary: 'MAX' }),
        search.createColumn({ name: 'formulanumeric', summary: 'AVG', formula: "CASE WHEN TO_CHAR({trandate}, 'YYYY') = '2021'  THEN {amount}/{quantity} END" }),
        search.createColumn({ name: 'formulanumeric', summary: 'AVG', formula: "CASE WHEN TO_CHAR({trandate}, 'YYYY') = '2022'  THEN {amount}/{quantity} END" }),
        search.createColumn({ name: 'formulanumeric', summary: 'AVG', formula: "CASE WHEN TO_CHAR({trandate}, 'YYYY') = '2023'  THEN {amount}/{quantity} END" }),
        search.createColumn({ name: 'formulanumeric', summary: 'AVG', formula: "CASE WHEN TO_CHAR({trandate}, 'YYYY') = '2024'  THEN {amount}/{quantity} END" }),
        search.createColumn({ name: 'formulanumeric', summary: 'AVG', formula: "CASE WHEN TO_CHAR({trandate}, 'YYYY') = '2025'  THEN {amount}/{quantity} END" }),
        search.createColumn({ name: 'formulanumeric', summary: 'AVG', formula: "CASE WHEN {trandate} BETWEEN ADD_MONTHS(SYSDATE, -6) AND SYSDATE THEN {amount}/{quantity} END" }),
        search.createColumn({ name: 'islotitem', join: 'item', summary: 'MAX' }) // << needed for type routing
      ]
    });

    const updates = [];
    let count = 0;

    s.run().each((r) => {
      count++;
      try {
        const itemId       = r.getValue({ name: 'internalid', join: 'item', summary: 'GROUP' });
        const isLotRaw     = r.getValue({ name: 'islotitem', join: 'item', summary: 'MAX' }); // boolean or 'T'/'F'
        const isLot        = (isLotRaw === true || String(isLotRaw).toUpperCase() === 'T');
        const itemTypeText = r.getText({ name: 'type', join: 'item', summary: 'GROUP' });
        const recType      = nsItemTypeFromText(itemTypeText, isLot);

        if (!itemId || !recType) {
          log.debug('Skipping row (unknown type)', { itemId, itemTypeText, isLot });
          return true;
        }

        // Item field values
        const itemFld2021 = toNumber(r.getValue({ name: 'custitem_average_cost_2021', join: 'item', summary: 'MAX' }));
        const itemFld2022 = toNumber(r.getValue({ name: 'custitem_average_cost_2022', join: 'item', summary: 'MAX' }));
        const itemFld2023 = toNumber(r.getValue({ name: 'custitem_average_cost_2023', join: 'item', summary: 'MAX' }));
        const itemFld2024 = toNumber(r.getValue({ name: 'custitem_average_cost_2024', join: 'item', summary: 'MAX' }));
        const itemFld2025 = toNumber(r.getValue({ name: 'custitem_average_cost_2025', join: 'item', summary: 'MAX' }));
        const itemFld52W  = toNumber(r.getValue({ name: 'custitem_average_cost_last_52_weeks', join: 'item', summary: 'MAX' }));

        // Calculated values
        const y2021 = toNumber(r.getValue(s.columns[8]));
        const y2022 = toNumber(r.getValue(s.columns[9]));
        const y2023 = toNumber(r.getValue(s.columns[10]));
        const y2024 = toNumber(r.getValue(s.columns[11]));
        const y2025 = toNumber(r.getValue(s.columns[12]));
        const w52   = toNumber(r.getValue(s.columns[13]));

        log.debug('Row values', {
          itemId, recType, itemTypeText, isLot,
          current: { itemFld2021, itemFld2022, itemFld2023, itemFld2024, itemFld2025, itemFld52W },
          calculated: { y2021, y2022, y2023, y2024, y2025, w52 }
        });

        const values = {};
        if (y2021 !== null && !approxEqual(y2021, itemFld2021)) values.custitem_average_cost_2021 = y2021;
        if (y2022 !== null && !approxEqual(y2022, itemFld2022)) values.custitem_average_cost_2022 = y2022;
        if (y2023 !== null && !approxEqual(y2023, itemFld2023)) values.custitem_average_cost_2023 = y2023;
        if (y2024 !== null && !approxEqual(y2024, itemFld2024)) values.custitem_average_cost_2024 = y2024;
        if (y2025 !== null && !approxEqual(y2025, itemFld2025)) values.custitem_average_cost_2025 = y2025;
        if (w52  !== null && !approxEqual(w52,  itemFld52W )) values.custitem_average_cost_last_52_weeks = w52;

        if (Object.keys(values).length > 0) {
          log.debug('Queued update', { itemId, recType, isLot, values });
          updates.push({ itemId: String(itemId), recType, values });
        } else {
          log.debug('No update needed', { itemId, recType, isLot });
        }

      } catch (e) {
        log.error('Row processing error', e);
      }
      return true;
    });

    log.audit('getInputData complete', { rowsProcessed: count, updatesQueued: updates.length });
    return updates;
  };

  // ---- map ----
  const map = (context) => {
    const data = JSON.parse(context.value || '{}');
    log.debug('map input', data);

    const { itemId, recType, values } = data || {};
    if (!itemId || !recType) {
      log.error('Invalid map data', data);
      return;
    }

    try {
      record.submitFields({
        type: recType, // will be LOT_NUMBERED_* when applicable
        id: itemId,
        values: values,
        options: { enableSourcing: false, ignoreMandatoryFields: true }
      });
      log.audit('Item updated', { itemId, recType, values });
      context.write(itemId, { updated: true, fields: Object.keys(values) });
    } catch (e) {
      log.error('submitFields failed', { itemId, recType, values, error: e });
      context.write(itemId, { updated: false, error: e.message });
    }
  };

  // ---- summarize ----
  const summarize = (summary) => {};

  return { getInputData, map, summarize };
});
