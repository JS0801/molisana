/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * Required Free-Form Text parameter: custscript_mi_btn_so_search
 * Authenticated internal deployment only; execute as the current role.
 */
define(['N/search', 'N/record', 'N/runtime', 'N/log'], (search, record, runtime, log) => {
    function loadSearch() {
        const id = runtime.getCurrentScript().getParameter({ name: 'custscript_mi_btn_so_search' });
        if (!id) throw Error('Set Suitelet parameter custscript_mi_btn_so_search to the original transaction saved search ID.');
        return search.load({ id });
    }
    function processOrder(soId) {
        // Recheck current eligibility and obtain only the matching line IDs.
        const lineSearch = loadSearch();
        lineSearch.filterExpression = [lineSearch.filterExpression, 'AND', ['internalid', 'anyof', soId]];
        lineSearch.columns = [search.createColumn({ name: 'line', sort: search.Sort.ASC })];
        const lineIds = new Set();
        const pages = lineSearch.runPaged({ pageSize: 1000 });
        for (const range of pages.pageRanges) {
            for (const row of pages.fetch({ index: range.index }).data) lineIds.add(String(row.getValue({ name: 'line' })));
        }
        if (!lineIds.size) return { updated: 0, message: 'No lines match the configured saved search on this sales order.' };
        const so = record.load({ type: record.Type.SALES_ORDER, id: soId, isDynamic: false });
        const lots = {};
        const reserved = {};
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
            const detail = so.getSublistSubrecord({ sublistId: 'item', fieldId: 'inventorydetail', line });
            if (detail.getLineCount({ sublistId: 'inventoryassignment' }) > 0) continue;
            // Include line index so lot availability is re-evaluated for each line.
            const lotKey = `${item}_${location}_${line}`;

            if (!lots[lotKey]) {
                const lotSearch = search.create({
                    type: 'inventorynumber',
                    filters: [
                        ['item', 'anyof', item],
                        'AND',
                        ['location', 'anyof', location],
                        'AND',
                        ['quantityavailable', 'greaterthanorequalto', quantity]
                    ],
                    columns: [
                        search.createColumn({
                            name: 'internalid',
                            sort: search.Sort.ASC
                        }),
                        'inventorynumber', 'quantityavailable'
                    ]
                });

                let selectedLotId = null;
                let selectedLotName = '';
                let preferredFound = false;

                // Check beyond the first 10 results for the preferred lot.
                const lotPages = lotSearch.runPaged({ pageSize: 1000 });

                for (const pageRange of lotPages.pageRanges) {
                    const page = lotPages.fetch({ index: pageRange.index });

                    for (const result of page.data) {
                        const lotId = result.getValue({ name: 'internalid' });
                        const reservationKey = `${item}_${location}_${lotId}`;
                        const available = Number(result.getValue({ name: 'quantityavailable' }) || 0);
                        if (available - (reserved[reservationKey] || 0) + 0.00000001 < quantity) continue;
                        const lotName = String(
                            result.getValue({ name: 'inventorynumber' }) || ''
                        );

                        // Keep the first available lot as the fallback.
                        if (!selectedLotId) {
                            selectedLotId = lotId;
                            selectedLotName = lotName;
                        }

                        if (lotName === String(item)) {
                            selectedLotId = lotId;
                            selectedLotName = lotName;
                            preferredFound = true;
                            break;
                        }
                    }

                    if (preferredFound) break;
                }

                if (!selectedLotId) {
                    throw Error(
                        `Item ${item}: no lot at location ${location} ` +
                        `has available quantity of at least ${quantity}.`
                    );
                }

                lots[lotKey] = selectedLotId;

                log.debug({
                    title: 'Lot selected',
                    details: {
                        item,
                        location,
                        quantity,
                        lotId: selectedLotId,
                        lotName: selectedLotName,
                        preferredFound
                    }
                });
            }
            detail.insertLine({ sublistId: 'inventoryassignment', line: 0 });
            detail.setSublistValue({ sublistId: 'inventoryassignment', fieldId: 'issueinventorynumber', line: 0, value: lots[lotKey] });
            detail.setSublistValue({ sublistId: 'inventoryassignment', fieldId: 'quantity', line: 0, value: quantity });
            const reservationKey = `${item}_${location}_${lots[lotKey]}`;
            reserved[reservationKey] = (reserved[reservationKey] || 0) + quantity;
            updated++;
            log.debug({ title: 'Inventory detail prepared', details: { soId, line: line + 1, item, lotId: lots[lotKey], quantity, location } });
        }
        if (updated) {
            so.save({ enableSourcing: false, ignoreMandatoryFields: false });
            log.audit({ title: 'SO inventory details saved', details: { soId, updated } });

        }
        return { updated, message: updated ? `Inventory detail added to ${updated} line(s).` : 'No changes: matching lines are closed, not fully committed, or already have inventory assignments.' };
    }
    function onRequest(context) {
        let result;
        try {
            if (context.request.method !== 'POST') throw Error('Use the sales order Assign Inventory Detail button. POST required.');
            const payload = JSON.parse(context.request.body || '{}');
            const soId = String(payload.soid || '');
            if (!/^[1-9][0-9]*$/.test(soId)) throw Error('A valid sales order internal ID is required.');
            result = { success: true, ...processOrder(soId) };
        } catch (error) {
            log.error({ title: 'SO inventory detail button failed', details: `${error.name}: ${error.message}` });
            result = { success: false, message: error.message || String(error) };
        }
        context.response.setHeader({ name: 'Content-Type', value: 'application/json; charset=utf-8' });
        context.response.write({ output: JSON.stringify(result) });
    }
    return { onRequest };
});
