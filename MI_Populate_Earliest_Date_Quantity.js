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
            columns: [
                search.createColumn({ name: 'item', summary: 'GROUP' })
            ]
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

        return data;
    }

    function map(context) {
        var row = JSON.parse(context.value);
        var itemId = row.itemId;
        if (!itemId) return;

        var itemType = getItemRecordType(itemId);
        if (!itemType) {
            log.error('Unable to find item type', itemId);
            return;
        }

        var earliest = getEarliestInbound(itemId);

        var currentValues;
        try {
            currentValues = search.lookupFields({
                type: itemType,
                id: itemId,
                columns: [FLD_DATE, FLD_QTY]
            });
        } catch (e) {
            log.error('lookupFields failed', { itemId: itemId, itemType: itemType, error: e });
            return;
        }

        var oldDate = currentValues[FLD_DATE];
        var oldQty = currentValues[FLD_QTY];

        var oldDateStr = '';
        if (oldDate) {
            try {
                oldDateStr = format.format({
                    value: oldDate,
                    type: format.Type.DATE
                });
            } catch (e) {
                oldDateStr = String(oldDate);
            }
        }

        oldQty = (oldQty === '' || oldQty == null) ? '' : Number(oldQty);

        if (!earliest) {
            if (!oldDate && (oldQty === '' || oldQty == null)) {
                return;
            }

            try {
                var clearObj = {};
                clearObj[FLD_DATE] = null;   // date field must be null
                clearObj[FLD_QTY] = '';

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
                log.error('Clear failed', { itemId: itemId, itemType: itemType, error: e });
            }

            return;
        }

        if (oldDateStr === earliest.dateStr && oldQty === earliest.qty) {
            return;
        }

        try {
            var updateObj = {};
            updateObj[FLD_DATE] = earliest.dateObj;
            updateObj[FLD_QTY] = earliest.qty;

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
                date: earliest.dateStr,
                qty: earliest.qty
            });
        } catch (e) {
            log.error('Update failed', { itemId: itemId, itemType: itemType, error: e });
        }
    }

    function getItemRecordType(itemId) {
        var result = search.create({
            type: 'item',
            filters: [['internalid', 'anyof', itemId]],
            columns: ['recordtype']
        }).run().getRange({ start: 0, end: 1 });

        if (!result || !result.length) {
            return null;
        }

        return result[0].recordType || result[0].getValue('recordtype');
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

        if (!results || !results.length) return null;

        var dateRaw = results[0].getValue('custrecord_port_eta');
        if (!dateRaw) return null;

        var dateObj, dateStr;
        try {
            dateObj = format.parse({
                value: dateRaw,
                type: format.Type.DATE
            });
            dateStr = format.format({
                value: dateObj,
                type: format.Type.DATE
            });
        } catch (e) {
            log.error('Date parse failed', { itemId: itemId, raw: dateRaw, error: e });
            return null;
        }

        return {
            dateObj: dateObj,
            dateStr: dateStr,
            qty: Number(results[0].getValue('quantityexpected')) || 0,
            inboundId: results[0].getValue('internalid')
        };
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