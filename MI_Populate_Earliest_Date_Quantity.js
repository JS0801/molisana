/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 */
define(['N/search', 'N/record', 'N/format', 'N/log'], function (search, record, format, log) {

    var FLD_DATE = 'custitem_mi_earliest_port_date';
    var FLD_QTY = 'custitem_mi_earliest_expected_quantity';

    function getInputData() {
        var itemsObj = {};
        var data = [];

        search.create({
            type: 'inboundshipment',
            filters: [['status', 'anyof', 'inTransit']],
            columns: [search.createColumn({ name: 'item', summary: 'GROUP' })]
        }).run().each(function (res) {
            var itemId = res.getValue({ name: 'item', summary: 'GROUP' });
            if (itemId) itemsObj[itemId] = true;
            return true;
        });

        search.create({
            type: 'item',
            filters: [
                [FLD_DATE, 'isnotempty', ''],
                'OR',
                [FLD_QTY, 'isnotempty', '']
            ],
            columns: ['internalid']
        }).run().each(function (res) {
            var itemId = res.getValue('internalid');
            if (itemId) itemsObj[itemId] = true;
            return true;
        });

        for (var itemId in itemsObj) {
            data.push({ itemId: itemId });
        }

        log.audit('getInputData count', data.length);
        return data;
    }

    function map(context) {
        var row = JSON.parse(context.value);
        var itemId = row.itemId;
        if (!itemId) return;

        log.debug('Map start', { itemId: itemId, contextValue: context.value });

        var itemType = getItemRecordType(itemId);
        log.debug('Item type result', { itemId: itemId, itemType: itemType });

        if (!itemType) {
            log.error('Unable to find item type', { itemId: itemId });
            return;
        }

        var earliest = getEarliestInbound(itemId);
        log.debug('Earliest inbound result', { itemId: itemId, earliest: earliest });

        var currentValues;
        try {
            currentValues = search.lookupFields({
                type: itemType,
                id: itemId,
                columns: [FLD_DATE, FLD_QTY]
            });
            log.debug('lookupFields result', { itemId: itemId, itemType: itemType, currentValues: currentValues });
        } catch (e) {
            log.error('lookupFields failed', { itemId: itemId, itemType: itemType, error: e });
            return;
        }

        var oldDateVal = getDateValueFromLookup(currentValues[FLD_DATE]);
        var oldQty = normalizeNumber(currentValues[FLD_QTY]);

        log.debug('Current normalized values', {
            itemId: itemId,
            rawDate: currentValues[FLD_DATE],
            normalizedDate: oldDateVal,
            rawQty: currentValues[FLD_QTY],
            normalizedQty: oldQty
        });

        var oldDateStr = '';
        if (oldDateVal) {
            try {
                oldDateStr = format.format({
                    value: oldDateVal,
                    type: format.Type.DATE
                });
            } catch (e) {
                log.error('Old date format failed', {
                    itemId: itemId,
                    oldDateVal: oldDateVal,
                    oldDateType: Object.prototype.toString.call(oldDateVal),
                    error: e
                });
            }
        }

        log.debug('Old date string', { itemId: itemId, oldDateStr: oldDateStr });

        // No inbound found -> clear values
        if (!earliest) {
            if (!oldDateVal && oldQty === null) {
                log.debug('Nothing to clear', { itemId: itemId });
                return;
            }

            try {
                var clearObj = {};
                clearObj[FLD_DATE] = null;
                clearObj[FLD_QTY] = null;

                log.debug('Submitting clear', {
                    itemId: itemId,
                    itemType: itemType,
                    values: clearObj
                });

                record.submitFields({
                    type: itemType,
                    id: itemId,
                    values: clearObj,
                    options: {
                        enableSourcing: false,
                        ignoreMandatoryFields: true
                    }
                });

                log.audit('Item cleared', {
                    itemId: itemId,
                    itemType: itemType
                });
            } catch (e) {
                log.error('Clear failed', {
                    itemId: itemId,
                    itemType: itemType,
                    error: e
                });
            }
            return;
        }

        if (oldDateStr === earliest.dateStr && oldQty === earliest.qty) {
            log.debug('No change needed', {
                itemId: itemId,
                oldDateStr: oldDateStr,
                newDateStr: earliest.dateStr,
                oldQty: oldQty,
                newQty: earliest.qty
            });
            return;
        }

        try {
            var updateObj = {};
            updateObj[FLD_DATE] = earliest.dateObj;
            updateObj[FLD_QTY] = earliest.qty;

            log.debug('Submitting update', {
                itemId: itemId,
                itemType: itemType,
                values: updateObj,
                dateType: Object.prototype.toString.call(earliest.dateObj),
                qtyType: typeof earliest.qty
            });

            record.submitFields({
                type: itemType,
                id: itemId,
                values: updateObj,
                options: {
                    enableSourcing: false,
                    ignoreMandatoryFields: true
                }
            });

            log.audit('Item updated', {
                itemId: itemId,
                itemType: itemType,
                inboundId: earliest.inboundId,
                oldDateStr: oldDateStr,
                newDateStr: earliest.dateStr,
                oldQty: oldQty,
                newQty: earliest.qty
            });
        } catch (e) {
            log.error('Update failed', {
                itemId: itemId,
                itemType: itemType,
                earliest: earliest,
                error: e
            });
        }
    }

    function getItemRecordType(itemId) {
        var results = search.create({
            type: search.Type.ITEM,
            filters: [['internalid', 'anyof', itemId]],
            columns: ['internalid']
        }).run().getRange({ start: 0, end: 1 });

        if (!results || !results.length) {
            return null;
        }

        return results[0].recordType || null;
    }

    function getEarliestInbound(itemId) {
        var results = search.create({
            type: 'inboundshipment',
            filters: [
                ['status', 'anyof', 'inTransit'],
                'AND',
                ['item', 'anyof', itemId]
            ],
            columns: [
                search.createColumn({ name: 'custrecord_port_eta', sort: search.Sort.ASC }),
                search.createColumn({ name: 'quantityexpected' }),
                search.createColumn({ name: 'internalid' })
            ]
        }).run().getRange({ start: 0, end: 1 });

        if (!results || !results.length) {
            return null;
        }

        var dateRaw = results[0].getValue('custrecord_port_eta');
        var qtyRaw = results[0].getValue('quantityexpected');
        var inboundId = results[0].getValue('internalid');

        log.debug('Raw inbound result', {
            itemId: itemId,
            inboundId: inboundId,
            dateRaw: dateRaw,
            qtyRaw: qtyRaw
        });

        if (!dateRaw) {
            return null;
        }

        var dateObj = getDateObject(dateRaw);
        if (!dateObj) {
            log.error('Could not convert inbound date', {
                itemId: itemId,
                inboundId: inboundId,
                rawDate: dateRaw
            });
            return null;
        }

        var dateStr = '';
        try {
            dateStr = format.format({
                value: dateObj,
                type: format.Type.DATE
            });
        } catch (e) {
            log.error('Date format failed', {
                itemId: itemId,
                inboundId: inboundId,
                dateObj: dateObj,
                error: e
            });
            return null;
        }

        return {
            dateObj: dateObj,
            dateStr: dateStr,
            qty: normalizeNumber(qtyRaw) || 0,
            inboundId: inboundId
        };
    }

    function getDateObject(value) {
        if (!value) return null;

        if (Object.prototype.toString.call(value) === '[object Date]') {
            return value;
        }

        try {
            return format.parse({
                value: value,
                type: format.Type.DATE
            });
        } catch (e) {
            return null;
        }
    }

    function getDateValueFromLookup(value) {
        if (!value) return null;

        if (Object.prototype.toString.call(value) === '[object Date]') {
            return value;
        }

        if (typeof value === 'string') {
            return getDateObject(value);
        }

        if (Array.isArray(value) && value.length && value[0] && value[0].value) {
            return getDateObject(value[0].value);
        }

        if (value.value) {
            return getDateObject(value.value);
        }

        return null;
    }

    function normalizeNumber(value) {
        if (value === '' || value === null || typeof value === 'undefined') {
            return null;
        }

        var n = Number(value);
        return isNaN(n) ? null : n;
    }

    function summarize(summary) {
        if (summary.inputSummary.error) {
            log.error('Input Error', summary.inputSummary.error);
        }

        summary.mapSummary.errors.iterator().each(function (key, error) {
            log.error('Map Error ' + key, error);
            return true;
        });
    }

    return {
        getInputData: getInputData,
        map: map,
        summarize: summarize
    };
});