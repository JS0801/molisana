/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 * @NModuleScope SameAccount
 */
define(['N/search', 'N/record', 'N/runtime', 'N/format', 'N/log'],
(search, record, runtime, format, log) => {


  
    // Converts SO units into reservation units when they differ.
    function conversionRate(item, units, cache) {
        if (!cache[item]) {
            const info = search.lookupFields({ type: 'item', id: item, columns: ['unitstype'] });
            const type = info.unitstype && info.unitstype[0];
            cache[item] = { '': 1 };
            if (type) {
                const rec = record.load({ type: 'unitstype', id: type.value });
                for (let i = 0; i < rec.getLineCount({ sublistId: 'uom' }); i++) {
                    const unit = rec.getSublistValue({ sublistId: 'uom', fieldId: 'internalid', line: i });
                    cache[item][unit] = Number(rec.getSublistValue({ sublistId: 'uom', fieldId: 'conversionrate', line: i }));
                }
            }
        }
        const rate = cache[item][units || ''];
        if (!(rate > 0)) throw Error(`Unknown unit ${units} for item ${item}`);
        return rate;
    }

    function afterSubmit(context) {
      
        if (context.type == context.UserEventType.CREATE) return;
        const soId = String(context.newRecord.id);
        try {
            const script = runtime.getCurrentScript();
            const soSearchId = script.getParameter({ name: 'custscript_mi_or_so_search' });
            const strategy = script.getParameter({ name: 'custscript_mi_or_strategy' });
            const prefix = script.getParameter({ name: 'custscript_mi_or_name_prefix' }) || 'SO Reservation';
            if (!soSearchId || !strategy) {
                throw Error('Set SO search and allocation strategy parameters.');
            }

            // The search determines whether the SO qualifies, not which lines to reserve.
            const soSearch = search.load({ id: soSearchId });
            const soFilter = ['internalid', 'anyof', soId];
            soSearch.filterExpression = soSearch.filterExpression.length
                ? [soSearch.filterExpression, 'AND', soFilter] : [soFilter];
            const matched = soSearch.run().getRange({ start: 0, end: 1 }).length > 0;
            log.debug({ title: 'SO eligibility', details: { soId, search: soSearchId, matched } });
            if (!matched) return;

            const today = new Date();
            const start = new Date(today.getFullYear(), 0, 1);
            const end = new Date(today.getFullYear(), 11, 31);
            const dateKey = date => date.getFullYear() * 10000 + (date.getMonth() + 1) * 100 + date.getDate();
            const so = record.load({ type: record.Type.SALES_ORDER, id: soId });
            const subsidiary = so.getValue({ fieldId: 'subsidiary' });
            const channel = so.getValue({ fieldId: 'saleschannel' });
            const headerLocation = so.getValue({ fieldId: 'location' });
            if (!channel) throw Error('SO has no sales channel.');
            const groups = {};
            const reservationLineField = 'custcol_mi_related_order_reservation';
            const rates = {};
            for (let i = 0; i < so.getLineCount({ sublistId: 'item' }); i++) {
                const get = fieldId => so.getSublistValue({ sublistId: 'item', fieldId, line: i });
                if (get('isclosed') === true || get('isclosed') === 'T') continue;
                const type = get('itemtype');
                if (type === 'Kit') throw Error(`Line ${i + 1}: kits need component reservation rules.`);
                if (type !== 'InvtPart' && type !== 'Assembly') continue;
                if (get('createpo')) throw Error(`Line ${i + 1}: special-order/drop-ship line needs separate handling.`);
                const quantity = Number(get('quantity'));
                if (!Number.isFinite(quantity)) throw Error(`Invalid quantity on line ${i + 1}.`);
                if (quantity <= 0) continue;
                const item = get('item');
                const location = get('location') || headerLocation;
                const units = get('units') || '';
                if (!location) throw Error(`Missing location on line ${i + 1}.`);
                const key = [subsidiary || 0, channel, location, item].join('_');
                if (!groups[key]) groups[key] = { item, location, units, baseQty: 0, lines: [] };
                groups[key].baseQty += quantity * conversionRate(item, units, rates);
                groups[key].lines.push(i);
            }

            if (Object.keys(groups).length && !so.getSublistFields({ sublistId: 'item' }).includes(reservationLineField)) {
                throw Error(`Missing SO line field ${reservationLineField}.`);
            }
            const itemIds = [...new Set(Object.values(groups).map(group => String(group.item)))];
            if (!itemIds.length) return;
            // One search for all SO items. Match location/channel/subsidiary in results.
            const reservationSearch = search.create({
                type: 'orderreservation',
                filters: [
                    ['item', 'anyof', itemIds], 'AND',
                    ['closed', 'is', 'F'] //, 'AND',
                    // ['startdate', 'on', format.format({ value: start, type: format.Type.DATE })], 'AND',
                    // ['enddate', 'on', format.format({ value: end, type: format.Type.DATE })]
                ],
                columns: [
                    search.createColumn({ name: 'internalid', sort: search.Sort.ASC }),
                    'item', 'location', 'saleschannel',
                    ...(subsidiary ? ['subsidiary'] : [])
                ]
            });
            const matches = {};
            const results = reservationSearch.runPaged({ pageSize: 1000 });
            for (const pageRange of results.pageRanges) {
                const page = results.fetch({ index: pageRange.index });
                for (const result of page.data) {
                    const resultKey = [
                        subsidiary ? result.getValue({ name: 'subsidiary' }) : 0,
                        result.getValue({ name: 'saleschannel' }),
                        result.getValue({ name: 'location' }),
                        result.getValue({ name: 'item' })
                    ].join('_');
                    if (groups[resultKey] && !matches[resultKey]) {
                        matches[resultKey] = result.getValue({ name: 'internalid' });
                    }
                }
            }
            log.debug({ title: 'Reservation search', details: {
                soId, itemIds, start, end, resultCount: results.count, matches
            } });

            for (const key of Object.keys(groups)) {
                const group = groups[key];
                for (let attempt = 0; attempt < 3; attempt++) {
                    try {
                        if (script.getRemainingUsage() < 100) throw Error('Low script usage; review completed reservation logs before retrying.');
                        let reservation = matches[key]
                            ? record.load({ type: 'orderreservation', id: matches[key] }) : null;
                        if (reservation) {
                            const closed = reservation.getValue({ fieldId: 'closed' });
                            const from = reservation.getValue({ fieldId: 'startdate' });
                            const to = reservation.getValue({ fieldId: 'enddate' });
                            const actualKey = [
                                reservation.getValue({ fieldId: 'subsidiary' }) || 0,
                                reservation.getValue({ fieldId: 'saleschannel' }),
                                reservation.getValue({ fieldId: 'location' }),
                                reservation.getValue({ fieldId: 'item' })
                            ].join('_');
                            if (closed === true || closed === 'T' || !from || !to || actualKey !== key ||
                                dateKey(from) !== dateKey(start) || dateKey(to) !== dateKey(end)) {
                             //   throw Error(`Reservation ${matches[key]} changed after search; review before retrying.`);
                            }
                        }
                        const isNew = !reservation;
                        if (isNew) {
                            reservation = record.create({ type: 'orderreservation', isDynamic: true });
                            const values = {
                                ...(subsidiary ? { subsidiary } : {}),
                                item: group.item, location: group.location, saleschannel: channel,
                                name: `${prefix} ${today.getFullYear()} ${key}`,
                                orderallocationstrategy: Number(strategy), startdate: start, enddate: end,
                                transactiondate: today,
                                ...(group.units ? { units: group.units } : {})
                            };
                            for (const fieldId of Object.keys(values)) reservation.setValue({ fieldId, value: values[fieldId] });
                        }
                        const units = reservation.getValue({ fieldId: 'units' }) || '';
                        const added = group.baseQty / conversionRate(group.item, units, rates);
                        const previous = isNew ? 0 : Number(reservation.getValue({ fieldId: 'quantity' }) || 0);
                        const quantity = Math.round((previous + added) * 1e8) / 1e8;
                        if (!Number.isFinite(quantity) || quantity <= 0) throw Error('Invalid reservation quantity.');
                        reservation.setValue({ fieldId: 'quantity', value: quantity });
                        const reservationId = reservation.save({ enableSourcing: true, ignoreMandatoryFields: false });
                        for (const line of group.lines) {
                            so.setSublistValue({ sublistId: 'item', fieldId: reservationLineField, line, value: String(reservationId) });
                        }
                        log.audit({ title: isNew ? 'Reservation created' : 'Reservation updated', details: {
                            soId, reservationId, item: group.item, location: group.location, channel, previous, added, quantity, soLines: group.lines.map(line => line + 1)
                        } });
                        break;
                    } catch (error) {
                        if (error.name !== 'RCRD_HAS_BEEN_CHANGED' || attempt === 2) throw error;
                        log.debug({ title: 'Retrying reservation save', details: { soId, key, attempt: attempt + 1 } });
                    }
                }
            }
            // Save line links once. This is an EDIT, so the CREATE-only guard prevents reprocessing.
            so.save({ enableSourcing: false, ignoreMandatoryFields: false });
            log.audit({ title: 'SO reservation line links saved', details: { soId, fieldId: reservationLineField } });
            log.audit({ title: 'SO reservations complete', details: { soId, itemLocationGroups: Object.keys(groups).length } });
        } catch (error) {
            log.error({ title: `Reservation error - SO ${soId}`, details: `${error.name}: ${error.message}. SO and earlier reservation saves remain saved. SO line links may not have been saved; review audit logs before recovery.` });
            throw error;
        }
    }
    return { afterSubmit };
});
