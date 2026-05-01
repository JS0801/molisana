/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 *
 * Fill Inventory Detail on Purchase Order lines from CSV.
 *
 * Updated for new Molisana file setup:
 * - Old file: PO row and Lot rows are separate
 * - New file: PO + Lot + Expiry + Qty are on same row
 *
 * This script supports both formats.
 */
define(['N/file', 'N/record', 'N/search'],
(file, record, search) => {

    const FILE_ID = 1291697;

    // PO line custom column / standard fields used for matching
    const FLD_LINE_ITEM      = 'item';
    const FLD_LINE_VENDORREF = 'custcol_mi_vendor_ref_number';
    const FLD_LINE_ISHIP     = 'custcol_mi_related_inbound';

    // =========================================================
    // 1) getInputData
    // =========================================================
    const getInputData = () => {
        log.audit('STAGE-1 ▶▶▶', 'getInputData START');

        let raw;
        try {
            raw = file.load({ id: FILE_ID }).getContents();
            log.audit('STAGE-1 file loaded', 'length=' + raw.length);
        } catch (e) {
            log.error('STAGE-1 file.load FAILED', e);
            return [];
        }

        const lines = raw.split(/\r?\n/);
        log.audit('STAGE-1 line count', lines.length);

        const rows = [];

        for (let i = 1; i < lines.length; i++) {
            if (!lines[i] || !lines[i].trim()) {
                log.debug('skip empty row', i);
                continue;
            }

            const c = parseCsv(lines[i]);

            const r = {
                shipment : clean(c[0]),
                vendorRef: clean(c[1]),
                po       : clean(c[5]),
                itemId   : clean(c[6]),
                lot      : clean(c[13]),
                exp      : clean(c[14]),
                qtyExp   : toNum(c[15]),
                qtyRec   : toNum(c[16])
            };

            rows.push(r);
            log.debug('row ' + i, JSON.stringify(r));
        }

        log.audit('STAGE-1 rows parsed', rows.length);

        /*
         * Group by shipment | vendorRef | itemId
         *
         * Old file:
         *   Row 1 = PO + Qty
         *   Row 2+ = Lot + Qty
         *
         * New Molisana file:
         *   Every row = PO + Lot + Qty
         *
         * So now we add PO when PO exists,
         * and also add Lot when Lot exists.
         */
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
                    shipment : curShip,
                    vendorRef: curVRef,
                    itemId   : r.itemId,
                    pos      : [],
                    poIndex  : {},
                    lots     : []
                };

                log.debug('NEW group', key);
            }

            const g = groups[key];

            const qty = r.qtyExp || r.qtyRec;

            // Add / combine PO qty
            // This works for:
            // - old file PO header rows
            // - new file lot rows where PO exists on every row
            if (r.po && qty > 0) {
                addPoQty(g, r.po, qty);
                log.debug('  + PO', r.po + ' qty=' + qty);
            }

            // Add lot qty
            // IMPORTANT: this is no longer else-if.
            // New file has PO and Lot on same row.
            if (r.lot && qty > 0) {
                g.lots.push({
                    lot: r.lot,
                    exp: r.exp,
                    qty: qty
                });

                log.debug('  + LOT', r.lot + ' qty=' + qty + ' exp=' + r.exp);
            }

            if (!r.po && !r.lot) {
                log.debug('  row has no PO and no Lot — ignored', JSON.stringify(r));
            }
        });

        // Remove internal poIndex before sending to map stage
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

        log.audit('STAGE-1 ◀◀◀', 'getInputData END, returning ' + list.length);
        return list;
    };

    // =========================================================
    // 2) map — split lots across POs, emit keyed by PO
    // =========================================================
    const map = (ctx) => {
        log.audit('STAGE-2 ▶▶▶', 'map START key=' + ctx.key);
        log.debug('STAGE-2 raw value', ctx.value);

        try {
            const g = JSON.parse(ctx.value);

            log.audit('STAGE-2 parsed',
                'ship=' + g.shipment + ' vRef=' + g.vendorRef +
                ' item=' + g.itemId +
                ' POs=' + (g.pos ? g.pos.length : 'UNDEF') +
                ' Lots=' + (g.lots ? g.lots.length : 'UNDEF'));

            if (!g.pos || !g.pos.length) {
                log.error('STAGE-2 NO POs', 'item=' + g.itemId + ' — nothing to emit');
                return;
            }

            if (!g.lots || !g.lots.length) {
                log.error('STAGE-2 NO LOTS', 'item=' + g.itemId + ' — nothing to emit');
                return;
            }

            const pos = g.pos.map(p => ({
                po: p.po,
                remaining: p.qty,
                lots: []
            }));

            const lots = g.lots.map(l => ({
                lot: l.lot,
                exp: l.exp,
                remaining: l.qty
            }));

            log.debug('STAGE-2 pos init', JSON.stringify(pos));
            log.debug('STAGE-2 lots init', JSON.stringify(lots));

            let pi = 0;

            for (const lot of lots) {
                log.debug('STAGE-2 process lot', lot.lot + ' remaining=' + lot.remaining);

                while (lot.remaining > 0 && pi < pos.length) {
                    if (pos[pi].remaining <= 0) {
                        log.debug('  PO full, advancing', pos[pi].po);
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
                        'item=' + g.itemId + ' ' + take +
                        ' of ' + lot.lot + ' → ' + pos[pi].po +
                        ' PO left=' + pos[pi].remaining +
                        ' lot left=' + lot.remaining);

                    if (pos[pi].remaining <= 0) {
                        pi++;
                    }
                }

                if (lot.remaining > 0) {
                    log.error('STAGE-2 LOT LEFTOVER',
                        'item=' + g.itemId +
                        ' lot=' + lot.lot +
                        ' leftover=' + lot.remaining +
                        ' PO capacity exhausted');
                }
            }

            let emitted = 0;

            pos.forEach((p, i) => {
                log.debug('STAGE-2 pos[' + i + '] final',
                    p.po + ' remaining=' + p.remaining +
                    ' lots=' + JSON.stringify(p.lots));

                if (!p.lots.length) {
                    log.audit('STAGE-2 skip — no lots assigned', p.po);
                    return;
                }

                const payload = {
                    shipment : g.shipment,
                    vendorRef: g.vendorRef,
                    itemId   : g.itemId,
                    po       : p.po,
                    lots     : p.lots
                };

                const keyOut = p.po;

                log.audit('STAGE-2 EMIT',
                    'key=' + keyOut + ' value=' + JSON.stringify(payload));

                ctx.write({
                    key: keyOut,
                    value: JSON.stringify(payload)
                });

                emitted++;
            });

            log.audit('STAGE-2 ◀◀◀',
                'map END item=' + g.itemId + ' emitted=' + emitted);

        } catch (e) {
            log.error('STAGE-2 map EXCEPTION', e.message + ' stack=' + e.stack);
        }
    };

    // =========================================================
    // 3) reduce — load PO, match line, write inventory detail
    // =========================================================
    const reduce = (ctx) => {
        log.audit('STAGE-3 ▶▶▶',
            'reduce START key(PO)=' + ctx.key +
            ' values=' + (ctx.values ? ctx.values.length : 'NONE'));

        try {
            const poNumber = ctx.key;

            ctx.values.forEach((v, i) => {
                log.debug('STAGE-3 incoming[' + i + ']', v);
            });

            log.debug('STAGE-3 searching PO', poNumber);

            const hits = search.create({
                type: 'purchaseorder',
                filters: [
                    ['tranid', 'is', poNumber], 'AND',
                    ['mainline', 'is', 'T']
                ],
                columns: ['internalid']
            }).run().getRange({
                start: 0,
                end: 5
            });

            log.audit('STAGE-3 PO search result', 'hits=' + hits.length);

            if (!hits.length) {
                log.error('STAGE-3 PO NOT FOUND', 'po=' + poNumber);
                return;
            }

            const poId = hits[0].getValue('internalid');

            log.audit('STAGE-3 PO id', poId);

            const rec = record.load({
                type: 'purchaseorder',
                id: poId,
                isDynamic: true
            });

            const lineCount = rec.getLineCount({
                sublistId: 'item'
            });

            log.audit('STAGE-3 PO line count', lineCount);

            for (let i = 0; i < lineCount; i++) {
                rec.selectLine({
                    sublistId: 'item',
                    line: i
                });

                const it = rec.getCurrentSublistValue({
                    sublistId: 'item',
                    fieldId: FLD_LINE_ITEM
                });

                const vr = rec.getCurrentSublistValue({
                    sublistId: 'item',
                    fieldId: FLD_LINE_VENDORREF
                });

                const isText = rec.getCurrentSublistText({
                    sublistId: 'item',
                    fieldId: FLD_LINE_ISHIP
                });

                const qty = rec.getCurrentSublistValue({
                    sublistId: 'item',
                    fieldId: 'quantity'
                });

                log.debug('STAGE-3 existing PO line ' + i,
                    'item=' + it +
                    ' vRef=' + vr +
                    ' iship=' + isText +
                    ' qty=' + qty);
            }

            ctx.values.forEach((v, pIdx) => {
                log.audit('STAGE-3 processing payload ' + pIdx, v);

                const p = JSON.parse(v);

                let idx = -1;

                for (let i = 0; i < lineCount; i++) {
                    rec.selectLine({
                        sublistId: 'item',
                        line: i
                    });

                    const lineItem = rec.getCurrentSublistValue({
                        sublistId: 'item',
                        fieldId: FLD_LINE_ITEM
                    });

                    const lineVRef = rec.getCurrentSublistValue({
                        sublistId: 'item',
                        fieldId: FLD_LINE_VENDORREF
                    });

                    const lineIShip = rec.getCurrentSublistText({
                        sublistId: 'item',
                        fieldId: FLD_LINE_ISHIP
                    });

                    log.debug('STAGE-3 compare PO line ' + i,
                        'item=' + lineItem + '/' + p.itemId +
                        ' vRef=' + lineVRef + '/' + p.vendorRef +
                        ' iship=' + lineIShip + '/' + p.shipment);

                    if (String(lineItem) === String(p.itemId) &&
                        String(lineVRef) === String(p.vendorRef) &&
                        String(lineIShip) === String(p.shipment)) {

                        idx = i;
                        log.audit('STAGE-3 MATCHED PO line', 'line=' + i);
                        break;
                    }
                }

if (idx === -1) {
    log.error('STAGE-3 NO LINE MATCH ON PO',
        'po=' + poNumber +
        ' item=' + p.itemId +
        ' vRef=' + p.vendorRef +
        ' iship=' + p.shipment);
    return;
}

// Build lot summary for audit log
var lotSummary = '';

for (var l = 0; l < p.lots.length; l++) {
    if (lotSummary) {
        lotSummary += ' | ';
    }

    lotSummary += 'Lot=' + p.lots[l].lot +
        ', Qty=' + p.lots[l].qty +
        ', Exp=' + p.lots[l].exp;
}

// Main log you asked for
log.audit('PO LINE LOT DETAIL',
    'PO=' + poNumber +
    ' | PO Internal ID=' + poId +
    ' | NS Line Index=' + idx +
    ' | UI Line No=' + (idx + 1) +
    ' | Item=' + p.itemId +
    ' | Vendor Ref=' + p.vendorRef +
    ' | Inbound Shipment=' + p.shipment +
    ' | Lots: ' + lotSummary
);

rec.selectLine({
    sublistId: 'item',
    line: idx
});

                const inv = rec.getCurrentSublistSubrecord({
                    sublistId: 'item',
                    fieldId: 'inventorydetail'
                });

                log.debug('STAGE-3 inv detail subrecord opened on PO line', idx);

                p.lots.forEach((lt, j) => {
                    log.debug('STAGE-3 adding lot ' + j,
                        lt.lot + ' qty=' + lt.qty + ' exp=' + lt.exp);

                    const assignCount = inv.getLineCount({
                        sublistId: 'inventoryassignment'
                    });

                    if (assignCount > j) {
                        inv.selectLine({
                            sublistId: 'inventoryassignment',
                            line: j
                        });
                    } else {
                        inv.selectNewLine({
                            sublistId: 'inventoryassignment'
                        });
                    }

                    inv.setCurrentSublistValue({
                        sublistId: 'inventoryassignment',
                        fieldId: 'receiptinventorynumber',
                        value: lt.lot
                    });

                    inv.setCurrentSublistValue({
                        sublistId: 'inventoryassignment',
                        fieldId: 'quantity',
                        value: lt.qty
                    });

                    if (lt.exp) {
                        const d = parseExpiryDate(lt.exp);

                        if (d) {
                            inv.setCurrentSublistValue({
                                sublistId: 'inventoryassignment',
                                fieldId: 'expirationdate',
                                value: d
                            });

                            log.debug('  exp set', d.toString());
                        } else {
                            log.error('  bad exp', lt.exp);
                        }
                    }

                    inv.commitLine({
                        sublistId: 'inventoryassignment'
                    });

                    log.debug('  committed assignment line');
                });

                rec.commitLine({
                    sublistId: 'item'
                });

                log.audit('STAGE-3 PO item line committed', 'line=' + idx);
            });

            log.debug('STAGE-3 about to save PO', poNumber);

            // const savedId = rec.save({
            //     ignoreMandatoryFields: true
            // });

            log.audit('STAGE-3 ◀◀◀ PO SAVED',
                'po=' + poNumber + ' id=' + savedId);

        } catch (e) {
            log.error('STAGE-3 reduce EXCEPTION', e.message + ' stack=' + e.stack);
        }
    };

    // =========================================================
    // 4) summarize
    // =========================================================
    const summarize = (s) => {
        log.audit('STAGE-4 ▶▶▶', 'summarize START');

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

        log.audit('STAGE-4 ◀◀◀', 'summarize END');
    };

    // =========================================================
    // helpers
    // =========================================================

    const addPoQty = (group, po, qty) => {
        if (!group.poIndex[po] && group.poIndex[po] !== 0) {
            group.poIndex[po] = group.pos.length;
            group.pos.push({
                po: po,
                qty: 0
            });
        }

        const idx = group.poIndex[po];
        group.pos[idx].qty += qty;
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

        // Supports:
        // 14/10/27
        // 14/10/2027
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

    return {
        getInputData,
        map,
        reduce,
        summarize
    };
});