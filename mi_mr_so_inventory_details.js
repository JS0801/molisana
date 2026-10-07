/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 * @NModuleScope SameAccount
 * Parameter: custscript_mi_inv_so_search (Free-Form Text, saved search ID).
 * Optional override; blank uses the supplied criteria embedded below.
 */
define(['N/search', 'N/record', 'N/runtime', 'N/log'], (search, record, runtime, log) => {
    function loadSearch() {
        const searchId = runtime.getCurrentScript().getParameter({ name: 'custscript_mi_inv_so_search' });
        if (searchId) return search.load({ id: searchId });
        return search.create({
            type: 'salesorder',
            settings: [{ name: 'consolidationtype', value: 'ACCTTYPE' }],
            filters: [
                ['type', 'anyof', 'SalesOrd'], 'AND', ['mainline', 'is', 'F'], 'AND',
                [['custbody_order_type_cp', 'anyof', '3'], 'OR', ['memomain', 'contains', 'Xmas'],
                    'OR', ['formulatext: {otherrefnum}', 'contains', 'CHRISTMAS'],
                    'OR', ['formulatext: {otherrefnum}', 'contains', 'Xmas'], 'OR', ['memomain', 'contains', 'CHRISTMAS']], 'AND',
                [["formulanumeric: CASE WHEN {item} LIKE '%XMAS%' OR {item} LIKE '%BACI%' THEN 1 ELSE 0 END", 'equalto', '0'], 'OR',
                    ["formulanumeric: CASE WHEN UPPER({item.displayname}) LIKE '%CHIOSTRO%' OR UPPER({item.displayname}) LIKE '%STREGA%' THEN 1 ELSE 0 END", 'equalto', '1']], 'AND',
                ['datecreated', 'within', 'thisyear'], 'AND',
                ['formulanumeric: {quantity} - NVL({quantitycommitted},0)', 'equalto', '0'], 'AND',
                ['inventorydetail.internalidnumber', 'isempty', ''], 'AND',
                ['status', 'anyof', 'SalesOrd:B'], 'AND', ['quantity', 'greaterthan', '0']
            ], columns: ['internalid', 'line']
        });
    }

    function getInputData() {
        const soSearch = loadSearch();
        // Collapse matching lines into one result per SO before map execution.
        soSearch.columns = [search.createColumn({ name: 'internalid', summary: search.Summary.GROUP, sort: search.Sort.ASC })];
        log.audit({ title: 'Inventory detail input search', details: 'Grouped by SO internal ID' });
        return soSearch;
    }

    function map(context) {
        const input = JSON.parse(context.value);
        const soId = input.values['GROUP(internalid)'] &&
            (input.values['GROUP(internalid)'].value || input.values['GROUP(internalid)']);
        if (!soId) throw Error(`Missing grouped SO internal ID: ${context.value}`);
        try {
            // Recheck current eligibility and obtain only the matching line IDs.
            const lineSearch = loadSearch();
            lineSearch.filterExpression = [lineSearch.filterExpression, 'AND', ['internalid', 'anyof', soId]];
            lineSearch.columns = [search.createColumn({ name: 'line', sort: search.Sort.ASC })];
            const lineIds = new Set();
            const pages = lineSearch.runPaged({ pageSize: 1000 });
            for (const range of pages.pageRanges) {
                for (const row of pages.fetch({ index: range.index }).data) lineIds.add(String(row.getValue({ name: 'line' })));
            }
            if (!lineIds.size) return;
            const so = record.load({ type: record.Type.SALES_ORDER, id: soId, isDynamic: false });
            const lots = {};
            let updated = 0;
            for (let line = 0; line < so.getLineCount({ sublistId: 'item' }); line++) {
                const get = fieldId => so.getSublistValue({ sublistId: 'item', fieldId, line });
                if (!lineIds.has(String(get('line')))) continue;
                const item = String(get('item'));
                const quantity = Number(get('quantity'));
                const committed = Number(get('quantitycommitted') || 0);
                if (get('isclosed') === true || get('isclosed') === 'T' || !Number.isFinite(quantity) || quantity <= 0 || Math.abs(quantity - committed) > 0.00000001) continue;
                const location = get('location') || so.getValue({ fieldId: 'location' });
                if (!location) throw Error(`Line ${line + 1}: no location.`);
                if (!lots[item]) {
                    // Lot NUMBER text is the item internal ID; resolve its own internal ID.
                    const matches = search.create({
                        type: 'inventorynumber',
                        filters: [['item', 'anyof', item], 'AND', ['inventorynumber', 'is', item]],
                        columns: ['internalid']
                    }).run().getRange({ start: 0, end: 2 });
                    if (matches.length !== 1) throw Error(`Item ${item}: expected one lot numbered "${item}", found ${matches.length}.`);
                    lots[item] = matches[0].getValue({ name: 'internalid' });
                }
                const detail = so.getSublistSubrecord({ sublistId: 'item', fieldId: 'inventorydetail', line });
                // Never replace existing assignments, including on restarted map executions.
                if (detail.getLineCount({ sublistId: 'inventoryassignment' }) > 0) {
                    log.debug({ title: 'Existing inventory detail skipped', details: { soId, line: line + 1, item } });
                    continue;
                }
                detail.insertLine({ sublistId: 'inventoryassignment', line: 0 });
                detail.setSublistValue({ sublistId: 'inventoryassignment', fieldId: 'issueinventorynumber', line: 0, value: lots[item] });
                detail.setSublistValue({ sublistId: 'inventoryassignment', fieldId: 'quantity', line: 0, value: quantity });
                updated++;
                log.debug({ title: 'Inventory detail prepared', details: { soId, line: line + 1, item, lotId: lots[item], quantity, location } });
            }
            if (updated) {
                so.save({ enableSourcing: false, ignoreMandatoryFields: false });
                log.audit({ title: 'SO inventory details saved', details: { soId, updated } });
                context.write({ key: String(soId), value: updated });
            }
        } catch (error) {
            log.error({ title: `SO inventory detail failed: ${soId}`, details: `${error.name}: ${error.message}` });
            throw error;
        }
    }

    function summarize(summary) {
        if (summary.inputSummary.error) log.error({ title: 'Input search failed', details: summary.inputSummary.error });
        summary.mapSummary.errors.iterator().each((key, error) => {
            log.error({ title: `Map failed: ${key}`, details: error });
            return true;
        });
        let orders = 0, lines = 0;
        summary.output.iterator().each((key, value) => { orders++; lines += Number(value); return true; });
        log.audit({ title: 'Inventory detail processing complete', details: { orders, lines, usage: summary.usage, seconds: summary.seconds, yields: summary.yields } });
    }
    return { getInputData, map, summarize };
});
