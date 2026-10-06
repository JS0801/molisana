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
        if (context.type !== context.UserEventType.CREATE) return;
        const soId = String(context.newRecord.id);
        try {
            const script = runtime.getCurrentScript();
            const soSearchId = script.getParameter({ name: 'custscript_mi_or_so_search' });
            const reservationSearchId = script.getParameter({ name: 'custscript_mi_or_res_search' });
            const strategy = script.getParameter({ name: 'custscript_mi_or_strategy' });
            const prefix = script.getParameter({ name: 'custscript_mi_or_name_prefix' }) || 'SO Reservation';
            const form = script.getParameter({ name: 'custscript_mi_or_form' });
            const firm = script.getParameter({ name: 'custscript_mi_or_firm' });
            const startParam = script.getParameter({ name: 'custscript_mi_or_start_date' });
            const endParam = script.getParameter({ name: 'custscript_mi_or_end_date' });
            if (!soSearchId || !reservationSearchId || !strategy) {
                throw Error('Set SO search, reservation search, allocation strategy parameters.');
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
            const start = startParam ? format.parse({ value: startParam, type: format.Type.DATE }) : new Date(today.getFullYear(), 0, 1);
            const end = endParam ? format.parse({ value: endParam, type: format.Type.DATE }) : new Date(today.getFullYear(), 11, 31);
            const dateKey = date => date.getFullYear() * 10000 + (date.getMonth() + 1) * 100 + date.getDate();
            if (!(start instanceof Date) || !(end instanceof Date) ||
                !(dateKey(start) <= dateKey(today) && dateKey(end) >= dateKey(today))) {
                throw Error('Reservation dates must include today. Leave both dates blank for the current calendar year.');
            }
            const so = record.load({ type: record.Type.SALES_ORDER, id: soId });
            const subsidiary = so.getValue({ fieldId: 'subsidiary' });
            const channel = so.getValue({ fieldId: 'saleschannel' });
            const headerLocation = so.getValue({ fieldId: 'location' });
            if (!channel) throw Error('SO has no sales channel.');
            const groups = {};
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
                if (!groups[key]) groups[key] = { item, location, units, baseQty: 0 };
                groups[key].baseQty += quantity * conversionRate(item, units, rates);
            }

            for (const key of Object.keys(groups)) {
                const group = groups[key];
                // Retry concurrent reservation saves.
                for (let attempt = 0; attempt < 3; attempt++) {
                    try {
                        const reservationSearch = search.load({ id: reservationSearchId });
                        const filters = [
                            ['mainline', 'is', 'T'], 'AND', ['item', 'anyof', group.item],
                            'AND', ['location', 'anyof', group.location],
                            'AND', ['saleschannel', 'anyof', channel]
                        ];
                        if (subsidiary) filters.push('AND', ['subsidiary', 'anyof', subsidiary]);
                        reservationSearch.filterExpression = reservationSearch.filterExpression.length
                            ? [reservationSearch.filterExpression, 'AND', filters] : filters;
                        reservationSearch.columns = [search.createColumn({ name: 'internalid', sort: search.Sort.ASC })];
                        const results = reservationSearch.run().getRange({ start: 0, end: 1000 });
                        if (results.length === 1000) throw Error('Too much reservation history for immediate processing.');
                        let reservation = null, lastId = 0;
                        for (const result of results) {
                            if (script.getRemainingUsage() < 100) throw Error('Low script usage; review completed reservation logs before retrying.');
                            const reservationId = result.getValue({ name: 'internalid' });
                            lastId = Math.max(lastId, Number(reservationId));
                            const candidate = record.load({ type: 'orderreservation', id: reservationId });
                            const closed = candidate.getValue({ fieldId: 'closed' });
                            const from = candidate.getValue({ fieldId: 'startdate' });
                            const to = candidate.getValue({ fieldId: 'enddate' });
                            if (!reservation && closed !== true && closed !== 'T' && from && to &&
                                dateKey(from) <= dateKey(today) && dateKey(to) >= dateKey(today)) reservation = candidate;
                        }
                        const isNew = !reservation;
                        if (isNew) {
                            reservation = record.create({ type: 'orderreservation', isDynamic: true });
                            const values = {
                                ...(form ? { customform: Number(form) } : {}),
                                ...(subsidiary ? { subsidiary } : {}),
                                item: group.item, location: group.location, saleschannel: channel,
                                name: `${prefix} ${today.getFullYear()} ${key}`,
                                orderallocationstrategy: Number(strategy), startdate: start, enddate: end,
                                transactiondate: today, commitmentfirm: firm === true || firm === 'T',
                                ...(group.units ? { units: group.units } : {}),
                                externalid: `MI_OR_${today.getFullYear()}_${key}_${lastId}`
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
                        log.audit({ title: isNew ? 'Reservation created' : 'Reservation updated', details: {
                            soId, reservationId, item: group.item, location: group.location, channel, previous, added, quantity
                        } });
                        break;
                    } catch (error) {
                        const conflict = /RCRD_HAS_BEEN_CHANGED|DUP.*(EXTERNAL|RCRD|RECORD)|DUPLICATE/.test(error.name || '') ||
                            /external id.*already|duplicate.*external id/i.test(error.message || '');
                        if (!conflict || attempt === 2) throw error;
                        log.debug({ title: 'Retrying reservation save', details: { soId, key, attempt: attempt + 1 } });
                    }
                }
            }
            log.audit({ title: 'SO reservations complete', details: { soId, itemLocationGroups: Object.keys(groups).length } });
        } catch (error) {
            log.error({ title: `Reservation error - SO ${soId}`, details: `${error.name}: ${error.message}. SO and earlier reservation saves remain saved.` });
            throw error;
        }
    }
    return { afterSubmit };
});
