/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 *
 * Test script: load one hardcoded Inbound Shipment and set inventory detail
 * on its item lines.
 */
define(['N/record'], (record) => {

    const INBOUND_SHIPMENT_ID = 12345; // TODO: replace with your IB internal ID

    const REC_INBOUND_SHIPMENT = 'inboundshipment';
    const SUBLIST_ITEMS = 'items';
    const SUBREC_INVENTORY_DETAIL = 'inventorydetail';
    const SUBLIST_ASSIGNMENT = 'inventoryassignment';

    const FLD_LINE_ITEM = 'shipmentitem';
    const FLD_LINE_QTY_CANDIDATES = [
        'quantityexpected',
        'quantityremaining',
        'quantity',
        'quantityreceived'
    ];

    /*
     * Empty array means apply to every inventory-detail-capable line.
     * Use zero-based NetSuite line indexes if you only want certain lines,
     * for example: [0, 2]
     */
    const LINE_INDEXES_TO_UPDATE = [];

    const INVENTORY_DETAIL_TO_SET = {
        lotNumber: 'TEST-LOT-001',
        quantity: null, // null = use the inbound shipment line quantity
        expirationDate: '12/31/2026',
        binNumber: null, // internal ID, or null
        inventoryStatus: null // internal ID, or null
    };

    const getInputData = () => {
        return [String(INBOUND_SHIPMENT_ID)];
    };

    const map = (context) => {
        context.write({
            key: context.value,
            value: context.value
        });
    };

    const reduce = (context) => {
        const inboundShipmentId = context.key;

        log.audit('START', 'Loading inbound shipment ' + inboundShipmentId);

        const ib = record.load({
            type: REC_INBOUND_SHIPMENT,
            id: inboundShipmentId,
            isDynamic: true
        });

        const lineCount = ib.getLineCount({
            sublistId: SUBLIST_ITEMS
        });

        let updated = 0;

        for (let line = 0; line < lineCount; line++) {
            if (LINE_INDEXES_TO_UPDATE.length && LINE_INDEXES_TO_UPDATE.indexOf(line) === -1) {
                continue;
            }

            try {
                if (setLineInventoryDetail(ib, line)) {
                    updated++;
                }
            } catch (e) {
                log.error('Line failed',
                    'line=' + line +
                    ' error=' + errorText(e) +
                    ' stack=' + (e && e.stack ? e.stack : ''));
            }
        }

        const savedId = ib.save({
            enableSourcing: false,
            ignoreMandatoryFields: true
        });

        log.audit('DONE', 'Inbound shipment saved=' + savedId + ' updatedLines=' + updated);
    };

    const setLineInventoryDetail = (ib, line) => {
        const itemId = safeGetSublistValue(ib, SUBLIST_ITEMS, FLD_LINE_ITEM, line);
        const lineQuantity = getLineQuantity(ib, line);
        const assignmentQuantity = INVENTORY_DETAIL_TO_SET.quantity == null
            ? lineQuantity
            : Number(INVENTORY_DETAIL_TO_SET.quantity);

        if (!assignmentQuantity || assignmentQuantity <= 0) {
            log.audit('Skipping line', 'line=' + line + ' item=' + itemId + ' no quantity found');
            return false;
        }

        ib.selectLine({
            sublistId: SUBLIST_ITEMS,
            line: line
        });

        const inventoryDetail = ib.getCurrentSublistSubrecord({
            sublistId: SUBLIST_ITEMS,
            fieldId: SUBREC_INVENTORY_DETAIL
        });

        clearInventoryAssignments(inventoryDetail);

        inventoryDetail.selectNewLine({
            sublistId: SUBLIST_ASSIGNMENT
        });

        inventoryDetail.setCurrentSublistValue({
            sublistId: SUBLIST_ASSIGNMENT,
            fieldId: 'receiptinventorynumber',
            value: INVENTORY_DETAIL_TO_SET.lotNumber
        });

        inventoryDetail.setCurrentSublistValue({
            sublistId: SUBLIST_ASSIGNMENT,
            fieldId: 'quantity',
            value: assignmentQuantity
        });

        if (INVENTORY_DETAIL_TO_SET.expirationDate) {
            inventoryDetail.setCurrentSublistValue({
                sublistId: SUBLIST_ASSIGNMENT,
                fieldId: 'expirationdate',
                value: parseDate(INVENTORY_DETAIL_TO_SET.expirationDate)
            });
        }

        if (INVENTORY_DETAIL_TO_SET.binNumber) {
            inventoryDetail.setCurrentSublistValue({
                sublistId: SUBLIST_ASSIGNMENT,
                fieldId: 'binnumber',
                value: INVENTORY_DETAIL_TO_SET.binNumber
            });
        }

        if (INVENTORY_DETAIL_TO_SET.inventoryStatus) {
            inventoryDetail.setCurrentSublistValue({
                sublistId: SUBLIST_ASSIGNMENT,
                fieldId: 'inventorystatus',
                value: INVENTORY_DETAIL_TO_SET.inventoryStatus
            });
        }

        inventoryDetail.commitLine({
            sublistId: SUBLIST_ASSIGNMENT
        });

        ib.commitLine({
            sublistId: SUBLIST_ITEMS
        });

        log.audit('Inventory detail set',
            'line=' + line +
            ' item=' + itemId +
            ' lot=' + INVENTORY_DETAIL_TO_SET.lotNumber +
            ' qty=' + assignmentQuantity);

        return true;
    };

    const clearInventoryAssignments = (inventoryDetail) => {
        const count = inventoryDetail.getLineCount({
            sublistId: SUBLIST_ASSIGNMENT
        });

        for (let line = count - 1; line >= 0; line--) {
            inventoryDetail.removeLine({
                sublistId: SUBLIST_ASSIGNMENT,
                line: line,
                ignoreRecalc: true
            });
        }
    };

    const getLineQuantity = (rec, line) => {
        for (let i = 0; i < FLD_LINE_QTY_CANDIDATES.length; i++) {
            const value = safeGetSublistValue(rec, SUBLIST_ITEMS, FLD_LINE_QTY_CANDIDATES[i], line);

            if (value !== null && value !== '' && value !== undefined && Number(value) > 0) {
                return Number(value);
            }
        }

        return 0;
    };

    const safeGetSublistValue = (rec, sublistId, fieldId, line) => {
        try {
            return rec.getSublistValue({
                sublistId: sublistId,
                fieldId: fieldId,
                line: line
            });
        } catch (e) {
            return null;
        }
    };

    const parseDate = (dateText) => {
        const parts = String(dateText).split('/');

        return new Date(
            Number(parts[2]),
            Number(parts[0]) - 1,
            Number(parts[1])
        );
    };

    const errorText = (e) => {
        if (e && e.name && e.message) {
            return e.name + ': ' + e.message;
        }

        return String(e || 'Unknown error');
    };

    return {
        getInputData: getInputData,
        map: map,
        reduce: reduce
    };
});
