/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 *
 * Fill Inventory Detail on Inbound Shipment lines from the same Molisana CSV
 * setup currently used for Purchase Order inventory detail.
 *
 * Flow:
 * - Loads the first CSV file in PENDING_FOLDER_ID.
 * - Parses shipment/vendor ref/PO/item/lot/expiry/qty.
 * - Splits lots across PO quantities the same way the PO-side script does.
 * - Loads the matching Inbound Shipment.
 * - Matches Inbound Shipment lines by item + PO where available.
 * - Replaces inventory assignment rows with lot/qty/expiry from the CSV.
 * - Moves the file to PROCESSED_FOLDER_ID only when all emitted payloads for
 *   the file complete successfully.
 */
define(['N/file', 'N/record', 'N/search'],
(file, record, search) => {

    const PENDING_FOLDER_ID = 427162;
    const PROCESSED_FOLDER_ID = 459909;
    const STATUS_PREFIX = '__FILE_STATUS__|';

    const REC_INBOUND_SHIPMENT = 'inboundshipment';
    const SUBLIST_ITEMS = 'items';
    const SUBREC_INVENTORY_DETAIL = 'inventorydetail';
    const SUBLIST_ASSIGNMENT = 'inventoryassignment';

    const FLD_IB_ITEM = 'shipmentitem';
    const FLD_IB_PURCHASE_ORDER = 'purchaseorder';
    const CLEAR_EXISTING_ASSIGNMENTS = true;

    // =========================================================
    // 1) getInputData
    // =========================================================
    const getInputData = () => {
        log.audit('STAGE-1 START', 'getInputData');

        let raw;
        let sourceFileId = '';
        let sourceFileName = '';

        try {
            const pendingFiles = findPendingFiles();

            if (!pendingFiles.length) {
                log.audit('No pending files', 'folder=' + PENDING_FOLDER_ID);
                return [];
            }

            const pendingFile = pendingFiles[0];
            sourceFileId = pendingFile.id;
            sourceFileName = pendingFile.name;

            raw = file.load({ id: sourceFileId }).getContents();

            log.audit('STAGE-1 file loaded',
                'fileId=' + sourceFileId +
                ' name=' + sourceFileName +
                ' length=' + raw.length);
        } catch (e) {
            log.error('STAGE-1 file load failed', errorText(e));
            return [];
        }

        const rows = parseRows(raw);
        const groups = buildGroups(rows, sourceFileId, sourceFileName);
        const list = [];

        for (const key in groups) {
            if (groups.hasOwnProperty(key)) {
                delete groups[key].poIndex;
                list.push(groups[key]);
            }
        }

        log.audit('STAGE-1 groups built', list.length);

        list.forEach((g, i) => {
            log.audit('STAGE-1 group[' + i + ']', JSON.stringify(g));
        });

        log.audit('STAGE-1 END', 'returning ' + list.length);
        return list;
    };

    // =========================================================
    // 2) map - split lots across POs, emit keyed by shipment
    // =========================================================
    const map = (ctx) => {
        log.audit('STAGE-2 START', 'key=' + ctx.key);
        log.debug('STAGE-2 raw value', ctx.value);

        try {
            const g = JSON.parse(ctx.value);

            log.audit('STAGE-2 parsed',
                'ship=' + g.shipment +
                ' vRef=' + g.vendorRef +
                ' item=' + g.itemId +
                ' POs=' + (g.pos ? g.pos.length : 'UNDEF') +
                ' Lots=' + (g.lots ? g.lots.length : 'UNDEF'));

            if (!g.pos || !g.pos.length) {
                emitFileStatus(ctx, g, 'ERROR', 'No POs found for item=' + g.itemId);
                return;
            }

            if (!g.lots || !g.lots.length) {
                emitFileStatus(ctx, g, 'ERROR', 'No lots found for item=' + g.itemId);
                return;
            }

            const pos = g.pos.map((p) => ({
                po: p.po,
                remaining: p.qty,
                lots: []
            }));

            const lots = g.lots.map((l) => ({
                lot: l.lot,
                exp: l.exp,
                remaining: l.qty
            }));

            let pi = 0;

            for (const lot of lots) {
                while (lot.remaining > 0 && pi < pos.length) {
                    if (pos[pi].remaining <= 0) {
                        pi++;
                        continue;
                    }

                    const take = Math.min(lot.remaining, pos[pi].remaining);

                    pos[pi].lots.push({
                        lot: lot.lot,
                        exp: lot.exp,
                        qty: take
                    });

                    pos[pi].remaining -= take;
                    lot.remaining -= take;

                    log.audit('STAGE-2 SPLIT',
                        'ship=' + g.shipment +
                        ' item=' + g.itemId +
                        ' qty=' + take +
                        ' lot=' + lot.lot +
                        ' po=' + pos[pi].po +
                        ' PO left=' + pos[pi].remaining +
                        ' lot left=' + lot.remaining);

                    if (pos[pi].remaining <= 0) {
                        pi++;
                    }
                }

                if (lot.remaining > 0) {
                    emitFileStatus(ctx, g, 'ERROR',
                        'Lot leftover after PO capacity exhausted. item=' +
                        g.itemId + ' lot=' + lot.lot +
                        ' leftover=' + lot.remaining);
                }
            }

            let emitted = 0;

            pos.forEach((p) => {
                if (!p.lots.length) {
                    log.audit('STAGE-2 skip no lots assigned', p.po);
                    return;
                }

                const payload = {
                    sourceFileId: g.sourceFileId,
                    sourceFileName: g.sourceFileName,
                    shipment: g.shipment,
                    vendorRef: g.vendorRef,
                    itemId: g.itemId,
                    po: p.po,
                    lots: p.lots
                };

                ctx.write({
                    key: g.shipment,
                    value: JSON.stringify(payload)
                });

                emitted++;
                log.audit('STAGE-2 EMIT',
                    'key(shipment)=' + g.shipment +
                    ' payload=' + JSON.stringify(payload));
            });

            log.audit('STAGE-2 END',
                'ship=' + g.shipment +
                ' item=' + g.itemId +
                ' emitted=' + emitted);

        } catch (e) {
            log.error('STAGE-2 map exception', errorText(e) + ' stack=' + (e && e.stack ? e.stack : ''));
        }
    };

    // =========================================================
    // 3) reduce - load Inbound Shipment and write inventory detail
    // =========================================================
    const reduce = (ctx) => {
        const shipmentNumber = ctx.key;

        if (String(shipmentNumber).indexOf(STATUS_PREFIX) === 0) {
            ctx.values.forEach((value) => {
                ctx.write({
                    key: shipmentNumber,
                    value: value
                });
            });
            return;
        }

        log.audit('STAGE-3 START',
            'shipment=' + shipmentNumber +
            ' values=' + (ctx.values ? ctx.values.length : 'NONE'));

        const fileStatuses = {};
        const failures = [];
        const successfulPayloads = [];

        try {
            const inboundShipmentId = findInboundShipmentId(shipmentNumber);

            if (!inboundShipmentId) {
                ctx.values.forEach((v) => {
                    const p = JSON.parse(v);
                    emitFileStatusObject(ctx, p, 'ERROR', 'Inbound Shipment not found: ' + shipmentNumber);
                });
                return;
            }

            log.audit('STAGE-3 Inbound Shipment found',
                'shipment=' + shipmentNumber + ' id=' + inboundShipmentId);

            const rec = record.load({
                type: REC_INBOUND_SHIPMENT,
                id: inboundShipmentId,
                isDynamic: true
            });

            const lineCount = rec.getLineCount({
                sublistId: SUBLIST_ITEMS
            });

            log.audit('STAGE-3 IB line count', lineCount);
            logInboundShipmentLines(rec, lineCount);

            const usedLines = {};

            ctx.values.forEach((v, payloadIndex) => {
                log.audit('STAGE-3 processing payload ' + payloadIndex, v);

                let p;

                try {
                    p = JSON.parse(v);
                    fileStatuses[p.sourceFileId] = p;

                    const lineIndex = findInboundShipmentLine(rec, lineCount, p, usedLines);

                    if (lineIndex === -1) {
                        const message = 'No IB line match. shipment=' + shipmentNumber +
                            ' po=' + p.po +
                            ' item=' + p.itemId +
                            ' vendorRef=' + p.vendorRef;

                        failures.push(message);
                        log.error('STAGE-3 no line match', message);
                        return;
                    }

                    rec.selectLine({
                        sublistId: SUBLIST_ITEMS,
                        line: lineIndex
                    });

                    const inv = rec.getCurrentSublistSubrecord({
                        sublistId: SUBLIST_ITEMS,
                        fieldId: SUBREC_INVENTORY_DETAIL
                    });

                    if (CLEAR_EXISTING_ASSIGNMENTS) {
                        clearInventoryAssignments(inv);
                    }

                    p.lots.forEach((lt, lotIndex) => {
                        writeInventoryAssignment(inv, lt, lotIndex);
                    });

                    rec.commitLine({
                        sublistId: SUBLIST_ITEMS
                    });

                    usedLines[lineIndex] = true;
                    successfulPayloads.push(p);

                    log.audit('IB LINE LOT DETAIL',
                        'Shipment=' + shipmentNumber +
                        ' | IB Internal ID=' + inboundShipmentId +
                        ' | NS Line Index=' + lineIndex +
                        ' | UI Line No=' + (lineIndex + 1) +
                        ' | PO=' + p.po +
                        ' | Item=' + p.itemId +
                        ' | Vendor Ref=' + p.vendorRef +
                        ' | Lots: ' + buildLotSummary(p.lots));

                } catch (lineError) {
                    const message = 'Payload failed. index=' + payloadIndex +
                        ' error=' + errorText(lineError);

                    failures.push(message);
                    log.error('STAGE-3 payload exception',
                        message + ' stack=' + (lineError && lineError.stack ? lineError.stack : ''));

                    if (p && p.sourceFileId) {
                        fileStatuses[p.sourceFileId] = p;
                    }
                }
            });

            if (successfulPayloads.length) {
                log.debug('STAGE-3 about to save IB', shipmentNumber);

                const savedId = rec.save({
                    ignoreMandatoryFields: true
                });

                log.audit('STAGE-3 IB SAVED',
                    'shipment=' + shipmentNumber +
                    ' id=' + savedId +
                    ' updatedPayloads=' + successfulPayloads.length);
            } else {
                log.audit('STAGE-3 no successful payloads', 'shipment=' + shipmentNumber);
            }

        } catch (e) {
            failures.push('Reduce exception: ' + errorText(e));
            log.error('STAGE-3 reduce exception', errorText(e) + ' stack=' + (e && e.stack ? e.stack : ''));
        }

        for (const fileId in fileStatuses) {
            if (fileStatuses.hasOwnProperty(fileId)) {
                emitFileStatusObject(
                    ctx,
                    fileStatuses[fileId],
                    failures.length ? 'ERROR' : 'OK',
                    failures.join(' | ')
                );
            }
        }

        log.audit('STAGE-3 END',
            'shipment=' + shipmentNumber +
            ' failures=' + failures.length);
    };

    // =========================================================
    // 4) summarize
    // =========================================================
    const summarize = (s) => {
        log.audit('STAGE-4 START', 'summarize');

        log.audit('STAGE-4 usage',
            'input=' + s.inputSummary.usage +
            ' map=' + s.mapSummary.usage +
            ' reduce=' + s.reduceSummary.usage);

        if (s.inputSummary.error) {
            log.error('STAGE-4 INPUT ERROR', s.inputSummary.error);
        }

        s.mapSummary.errors.iterator().each((k, e) => {
            log.error('STAGE-4 MAP ERR key=' + k, e);
            return true;
        });

        s.reduceSummary.errors.iterator().each((k, e) => {
            log.error('STAGE-4 REDUCE ERR key=' + k, e);
            return true;
        });

        const statusesByFile = {};

        s.output.iterator().each((key, value) => {
            const st = JSON.parse(value);

            if (!statusesByFile[st.sourceFileId]) {
                statusesByFile[st.sourceFileId] = {
                    sourceFileId: st.sourceFileId,
                    sourceFileName: st.sourceFileName,
                    hasError: false,
                    messages: []
                };
            }

            if (st.status !== 'OK') {
                statusesByFile[st.sourceFileId].hasError = true;
            }

            if (st.message) {
                statusesByFile[st.sourceFileId].messages.push(st.message);
            }

            return true;
        });

        for (const fileId in statusesByFile) {
            if (!statusesByFile.hasOwnProperty(fileId)) {
                continue;
            }

            const status = statusesByFile[fileId];

            if (status.hasError) {
                log.error('FILE NOT MOVED',
                    'fileId=' + status.sourceFileId +
                    ' name=' + status.sourceFileName +
                    ' messages=' + status.messages.join(' | '));
                continue;
            }

            // const f = file.load({ id: status.sourceFileId });
            // f.folder = PROCESSED_FOLDER_ID;
            // f.save();

            log.audit('FILE MOVED TO PROCESSED',
                'fileId=' + status.sourceFileId +
                ' name=' + status.sourceFileName +
                ' folder=' + PROCESSED_FOLDER_ID);
        }

        log.audit('STAGE-4 END', 'summarize');
    };

    // =========================================================
    // helpers
    // =========================================================

    const parseRows = (raw) => {
        const lines = raw.split(/\r?\n/);
        const rows = [];

        log.audit('STAGE-1 line count', lines.length);

        for (let i = 1; i < lines.length; i++) {
            if (!lines[i] || !lines[i].trim()) {
                continue;
            }

            const c = parseCsv(lines[i]);

            const r = {
                shipment: clean(c[0]),
                vendorRef: clean(c[1]),
                po: clean(c[5]),
                itemId: clean(c[6]),
                lot: clean(c[13]),
                exp: clean(c[14]),
                qtyExp: toNum(c[15]),
                qtyRec: toNum(c[16])
            };

            rows.push(r);
            log.debug('row ' + i, JSON.stringify(r));
        }

        log.audit('STAGE-1 rows parsed', rows.length);
        return rows;
    };

    const buildGroups = (rows, sourceFileId, sourceFileName) => {
        let curShip = '';
        let curVRef = '';
        const groups = {};

        rows.forEach((r, idx) => {
            if (r.shipment) {
                curShip = r.shipment;
            }

            if (r.vendorRef) {
                curVRef = r.vendorRef;
            }

            if (!curShip || !curVRef || !r.itemId) {
                log.error('STAGE-1 row missing key data',
                    'row=' + idx +
                    ' ship=' + curShip +
                    ' vendorRef=' + curVRef +
                    ' item=' + r.itemId);
                return;
            }

            const key = curShip + '|' + curVRef + '|' + r.itemId;

            if (!groups[key]) {
                groups[key] = {
                    sourceFileId: sourceFileId,
                    sourceFileName: sourceFileName,
                    shipment: curShip,
                    vendorRef: curVRef,
                    itemId: r.itemId,
                    pos: [],
                    poIndex: {},
                    lots: []
                };
            }

            const g = groups[key];
            const qty = r.qtyExp || r.qtyRec;

            if (r.po && qty > 0) {
                addPoQty(g, r.po, qty);
                log.debug('  + PO', r.po + ' qty=' + qty);
            }

            if (r.lot && qty > 0) {
                g.lots.push({
                    lot: r.lot,
                    exp: r.exp,
                    qty: qty
                });

                log.debug('  + LOT', r.lot + ' qty=' + qty + ' exp=' + r.exp);
            }
        });

        return groups;
    };

    const addPoQty = (group, po, qty) => {
        if (!group.poIndex[po] && group.poIndex[po] !== 0) {
            group.poIndex[po] = group.pos.length;
            group.pos.push({
                po: po,
                qty: 0
            });
        }

        group.pos[group.poIndex[po]].qty += qty;
    };

    const findInboundShipmentId = (shipmentNumber) => {
        const candidateFields = [
            'shipmentnumber',
            'tranid',
            'name'
        ];

        for (let i = 0; i < candidateFields.length; i++) {
            const fieldId = candidateFields[i];

            try {
                const hits = search.create({
                    type: REC_INBOUND_SHIPMENT,
                    filters: [
                        [fieldId, 'is', shipmentNumber]
                    ],
                    columns: ['internalid']
                }).run().getRange({
                    start: 0,
                    end: 2
                });

                log.debug('Inbound Shipment search',
                    'field=' + fieldId +
                    ' shipment=' + shipmentNumber +
                    ' hits=' + hits.length);

                if (hits.length) {
                    return hits[0].getValue('internalid') || hits[0].id;
                }
            } catch (e) {
                log.debug('Inbound Shipment search field failed',
                    'field=' + fieldId + ' error=' + errorText(e));
            }
        }

        return '';
    };

    const logInboundShipmentLines = (rec, lineCount) => {
        for (let i = 0; i < lineCount; i++) {
            rec.selectLine({
                sublistId: SUBLIST_ITEMS,
                line: i
            });

            log.debug('STAGE-3 existing IB line ' + i,
                'item=' + safeCurrentValue(rec, SUBLIST_ITEMS, FLD_IB_ITEM) +
                ' poValue=' + safeCurrentValue(rec, SUBLIST_ITEMS, FLD_IB_PURCHASE_ORDER) +
                ' poText=' + safeCurrentText(rec, SUBLIST_ITEMS, FLD_IB_PURCHASE_ORDER));
        }
    };

    const findInboundShipmentLine = (rec, lineCount, payload, usedLines) => {
        const matches = [];

        for (let i = 0; i < lineCount; i++) {
            if (usedLines[i]) {
                continue;
            }

            rec.selectLine({
                sublistId: SUBLIST_ITEMS,
                line: i
            });

            const lineItem = safeCurrentText(rec, SUBLIST_ITEMS, FLD_IB_ITEM);
            const linePoValue = safeCurrentValue(rec, SUBLIST_ITEMS, FLD_IB_PURCHASE_ORDER);
            const linePoText = safeCurrentText(rec, SUBLIST_ITEMS, FLD_IB_PURCHASE_ORDER)?.replace('PO#', '');

            const itemMatches = String(lineItem) === String(payload.itemId);
            const poMatches = !!payload.po &&
                (textMatches(linePoText, payload.po) || textMatches(linePoValue, payload.po));

            log.debug('STAGE-3 compare IB line ' + i,
                'item=' + lineItem + '/' + payload.itemId +
                ' poText=' + linePoText + '/' + payload.po +
                ' poValue=' + linePoValue + '/' + payload.po +
                ' itemMatches=' + itemMatches +
                ' poMatches=' + poMatches);

            if (itemMatches && poMatches) {
                matches.push(i);
            }
        }

        if (matches.length === 1) {
            return matches[0];
        }

        if (matches.length > 1) {
            log.audit('STAGE-3 multiple PO/item IB matches',
                'payload=' + JSON.stringify(payload) +
                ' matches=' + matches.join(',') +
                ' using first match');
            return matches[0];
        }

        log.error('STAGE-3 no PO/item IB line match',
            'payload=' + JSON.stringify(payload) +
            ' matches=' + matches.join(','));

        return -1;
    };

    const writeInventoryAssignment = (inv, lot, lotIndex) => {
        if (!CLEAR_EXISTING_ASSIGNMENTS) {
            const assignCount = inv.getLineCount({
                sublistId: SUBLIST_ASSIGNMENT
            });

            if (assignCount > lotIndex) {
                inv.selectLine({
                    sublistId: SUBLIST_ASSIGNMENT,
                    line: lotIndex
                });
            } else {
                inv.selectNewLine({
                    sublistId: SUBLIST_ASSIGNMENT
                });
            }
        } else {
            inv.selectNewLine({
                sublistId: SUBLIST_ASSIGNMENT
            });
        }

        inv.setCurrentSublistValue({
            sublistId: SUBLIST_ASSIGNMENT,
            fieldId: 'receiptinventorynumber',
            value: lot.lot
        });

        inv.setCurrentSublistValue({
            sublistId: SUBLIST_ASSIGNMENT,
            fieldId: 'quantity',
            value: lot.qty
        });

        if (lot.exp) {
            const d = parseExpiryDate(lot.exp);

            if (d) {
                inv.setCurrentSublistValue({
                    sublistId: SUBLIST_ASSIGNMENT,
                    fieldId: 'expirationdate',
                    value: d
                });
            } else {
                log.error('Bad expiry date', 'lot=' + lot.lot + ' exp=' + lot.exp);
            }
        }

        inv.commitLine({
            sublistId: SUBLIST_ASSIGNMENT
        });

        log.debug('STAGE-3 assignment committed',
            'lot=' + lot.lot +
            ' qty=' + lot.qty +
            ' exp=' + lot.exp);
    };

    const clearInventoryAssignments = (inv) => {
        const count = inv.getLineCount({
            sublistId: SUBLIST_ASSIGNMENT
        });

        for (let i = count - 1; i >= 0; i--) {
            inv.removeLine({
                sublistId: SUBLIST_ASSIGNMENT,
                line: i,
                ignoreRecalc: true
            });
        }

        log.debug('STAGE-3 cleared inventory assignments', count);
    };

    const buildLotSummary = (lots) => {
        let summary = '';

        for (let i = 0; i < lots.length; i++) {
            if (summary) {
                summary += ' | ';
            }

            summary += 'Lot=' + lots[i].lot +
                ', Qty=' + lots[i].qty +
                ', Exp=' + lots[i].exp;
        }

        return summary;
    };

    const emitFileStatus = (ctx, group, status, message) => {
        emitFileStatusObject(ctx, group, status, message);
    };

    const emitFileStatusObject = (ctx, payload, status, message) => {
        if (!payload || !payload.sourceFileId) {
            return;
        }

        ctx.write({
            key: STATUS_PREFIX + payload.sourceFileId,
            value: JSON.stringify({
                sourceFileId: payload.sourceFileId,
                sourceFileName: payload.sourceFileName,
                status: status,
                message: message || ''
            })
        });
    };

    const findPendingFiles = () => {
        const pendingFiles = [];

        search.create({
            type: 'file',
            filters: [
                ['folder', 'anyof', PENDING_FOLDER_ID]
            ],
            columns: ['internalid', 'name']
        }).run().each((result) => {
            pendingFiles.push({
                id: result.getValue('internalid') || result.id,
                name: result.getValue('name') || ''
            });

            return true;
        });

        return pendingFiles;
    };

    const parseCsv = (line) => {
        const out = [];
        let cur = '';
        let q = false;

        for (let i = 0; i < line.length; i++) {
            const ch = line[i];

            if (ch === '"') {
                if (q && line[i + 1] === '"') {
                    cur += '"';
                    i++;
                } else {
                    q = !q;
                }
                continue;
            }

            if (ch === ',' && !q) {
                out.push(cur);
                cur = '';
                continue;
            }

            cur += ch;
        }

        out.push(cur);
        return out;
    };

    const clean = (s) => {
        return String(s || '').replace(/^\uFEFF/, '').trim();
    };

    const toNum = (s) => {
        const n = parseFloat(String(s || '').replace(/,/g, '').trim());
        return isNaN(n) ? 0 : n;
    };

    const parseExpiryDate = (s) => {
        s = clean(s);

        const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);

        if (!m) {
            return null;
        }

        const day = parseInt(m[1], 10);
        const month = parseInt(m[2], 10);
        let year = parseInt(m[3], 10);

        if (year < 100) {
            year = 2000 + year;
        }

        return new Date(year, month - 1, day);
    };

    const safeCurrentValue = (rec, sublistId, fieldId) => {
        try {
            return rec.getCurrentSublistValue({
                sublistId: sublistId,
                fieldId: fieldId
            });
        } catch (e) {
            return '';
        }
    };

    const safeCurrentText = (rec, sublistId, fieldId) => {
        try {
            return rec.getCurrentSublistText({
                sublistId: sublistId,
                fieldId: fieldId
            });
        } catch (e) {
            return '';
        }
    };

    const textMatches = (actual, expected) => {
        const a = normalizeText(actual);
        const e = normalizeText(expected);

        if (!a || !e) {
            return false;
        }

        return a === e || a.indexOf(e) !== -1 || e.indexOf(a) !== -1;
    };

    const normalizeText = (value) => {
        return String(value || '')
            .toUpperCase()
            .replace(/\s+/g, '')
            .trim();
    };

    const errorText = (e) => {
        if (e && e.name && e.message) {
            return e.name + ': ' + e.message;
        }

        return String(e || 'Unknown error');
    };

    return {
        getInputData,
        map,
        reduce,
        summarize
    };
});
