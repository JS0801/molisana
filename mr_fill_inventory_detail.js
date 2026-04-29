/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 *
 * Fill Inventory Detail on Purchase Order lines from CSV.
 *
 * Flow:
 *   Stage 1 (getInputData)
 *     - Load CSV file by ID
 *     - Parse rows; group by shipment | vendorRef | itemId
 *     - Each group has list of POs (with expected qty) and list of Lots (with qty)
 *
 *   Stage 2 (map)
 *     - For each group, greedily split lot quantities across POs in order
 *     - Emit one payload per (PO, item) combination, keyed by PO number
 *
 *   Stage 3 (reduce)
 *     - Key = PO number; values = all item payloads for that PO
 *     - Load PO, find correct line by item + vendor ref + inbound shipment
 *     - Open inventorydetail subrecord on PO line, add lot assignments, save
 *
 *   Stage 4 (summarize)
 *     - Log usage and any errors
 *
 * Heavy logging — every meaningful step is traced.
 */
define(['N/file', 'N/record', 'N/search'],
(file, record, search) => {

    const FILE_ID = 1278609;

    // PO line custom column / standard fields used for matching
    const FLD_LINE_ITEM      = 'item';
    const FLD_LINE_VENDORREF = 'custcol_mi_vendor_ref_number'; // adjust if different on PO
    const FLD_LINE_ISHIP     = 'inboundshipment';              // standard NS field on PO line

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

        // Parse rows
        const rows = [];
        for (let i = 1; i < lines.length; i++) {
            if (!lines[i].trim()) { log.debug('skip empty row', i); continue; }
            const c = parseCsv(lines[i]);
            const r = {
                shipment : c[0],
                vendorRef: c[1],
                po       : c[5],
                itemId   : c[6],
                lot      : c[13],
                exp      : c[14],
                qtyExp   : toNum(c[15]),
                qtyRec   : toNum(c[16])
            };
            rows.push(r);
            log.debug('row ' + i, JSON.stringify(r));
        }
        log.audit('STAGE-1 rows parsed', rows.length);

        // Group by shipment | vendorRef | itemId
        let curShip = '', curVRef = '';
        const groups = {};

        rows.forEach((r, idx) => {
            if (r.shipment)  curShip = r.shipment;
            if (r.vendorRef) curVRef = r.vendorRef;

            const key = curShip + '|' + curVRef + '|' + r.itemId;
            log.debug('group key for row ' + idx, key);

            if (!groups[key]) {
                groups[key] = {
                    shipment : curShip,
                    vendorRef: curVRef,
                    itemId   : r.itemId,
                    pos      : [],
                    lots     : []
                };
                log.debug('NEW group', key);
            }

            if (r.po) {
                groups[key].pos.push({ po: r.po, qty: r.qtyExp });
                log.debug('  + PO', r.po + ' qty=' + r.qtyExp);
            } else if (r.lot) {
                // Lot qty in this CSV lives in c[15] (Qty Expected column).
                // Fall back to c[16] in case future exports flip it.
                const lotQty = r.qtyExp || r.qtyRec;
                groups[key].lots.push({ lot: r.lot, exp: r.exp, qty: lotQty });
                log.debug('  + LOT', r.lot + ' qty=' + lotQty);
            } else {
                log.debug('  row has no PO and no Lot — ignored', JSON.stringify(r));
            }
        });

        const list = Object.values(groups);
        log.audit('STAGE-1 groups built', list.length);

        list.forEach((g, i) => {
            log.audit('STAGE-1 group[' + i + ']', JSON.stringify(g));
        });

        log.audit('STAGE-1 ◀◀◀', 'getInputData END, returning ' + list.length);
        return list;
    };

    // =========================================================
    // 2) map  —  split lots across POs, emit keyed by PO
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

            // Working copies
            const pos = g.pos.map(p =>
                ({ po: p.po, remaining: p.qty, lots: [] }));
            const lots = g.lots.map(l =>
                ({ lot: l.lot, exp: l.exp, remaining: l.qty }));

            log.debug('STAGE-2 pos init', JSON.stringify(pos));
            log.debug('STAGE-2 lots init', JSON.stringify(lots));

            // Greedy split: walk lots, fill PO buckets in order
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
                    pos[pi].lots.push({ lot: lot.lot, exp: lot.exp, qty: take });
                    pos[pi].remaining -= take;
                    lot.remaining -= take;

                    log.audit('STAGE-2 SPLIT',
                        'item=' + g.itemId + ' ' + take +
                        ' of ' + lot.lot + ' → ' + pos[pi].po +
                        ' (PO left=' + pos[pi].remaining +
                        ', lot left=' + lot.remaining + ')');

                    if (pos[pi].remaining <= 0) pi++;
                }
                if (lot.remaining > 0) {
                    log.error('STAGE-2 LOT LEFTOVER',
                        'item=' + g.itemId + ' lot=' + lot.lot +
                        ' leftover=' + lot.remaining +
                        ' (PO capacity exhausted)');
                }
            }

            // Emit one payload per PO, keyed by PO number
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

                const keyOut = p.po; // group reduce by PO number
                log.audit('STAGE-2 EMIT', 'key=' + keyOut +
                    ' value=' + JSON.stringify(payload));
                ctx.write({ key: keyOut, value: JSON.stringify(payload) });
                emitted++;
            });

            log.audit('STAGE-2 ◀◀◀', 'map END item=' + g.itemId +
                ' emitted=' + emitted);

        } catch (e) {
            log.error('STAGE-2 map EXCEPTION', e.message + ' stack=' + e.stack);
        }
    };

    // =========================================================
    // 3) reduce  —  load PO, match line by item+vRef+iship, write inv detail
    // =========================================================
    const reduce = (ctx) => {
        log.audit('STAGE-3 ▶▶▶', 'reduce START key(PO)=' + ctx.key +
            ' values=' + (ctx.values ? ctx.values.length : 'NONE'));

        try {
            const poNumber = ctx.key;

            ctx.values.forEach((v, i) => {
                log.debug('STAGE-3 incoming[' + i + ']', v);
            });

            // ---- find PO by tranid ----
            log.debug('STAGE-3 searching PO', poNumber);
            const hits = search.create({
                type: 'purchaseorder',
                filters: [
                    ['tranid', 'is', poNumber], 'AND',
                    ['mainline', 'is', 'T']
                ],
                columns: ['internalid']
            }).run().getRange({ start: 0, end: 5 });

            log.audit('STAGE-3 PO search result', 'hits=' + hits.length);

            if (!hits.length) {
                log.error('STAGE-3 PO NOT FOUND', 'po=' + poNumber);
                return;
            }

            const poId = hits[0].getValue('internalid');
            log.audit('STAGE-3 PO id', poId);

            const rec = record.load({
                type: 'purchaseorder', id: poId, isDynamic: true
            });
            const lineCount = rec.getLineCount({ sublistId: 'item' });
            log.audit('STAGE-3 PO line count', lineCount);

            // ---- log every existing PO line for traceability ----
            for (let i = 0; i < lineCount; i++) {
                rec.selectLine({ sublistId: 'item', line: i });
                const it = rec.getCurrentSublistValue({
                    sublistId: 'item', fieldId: FLD_LINE_ITEM });
                const vr = rec.getCurrentSublistValue({
                    sublistId: 'item', fieldId: FLD_LINE_VENDORREF });
                const isText = rec.getCurrentSublistText({
                    sublistId: 'item', fieldId: FLD_LINE_ISHIP });
                const qty = rec.getCurrentSublistValue({
                    sublistId: 'item', fieldId: 'quantity' });
                log.debug('STAGE-3 existing PO line ' + i,
                    'item=' + it + ' vRef=' + vr +
                    ' iship=' + isText + ' qty=' + qty);
            }

            // ---- process each payload (one item per payload) ----
            ctx.values.forEach((v, pIdx) => {
                log.audit('STAGE-3 processing payload ' + pIdx, v);
                const p = JSON.parse(v);

                // match line by item + vendor ref + inbound shipment
                let idx = -1;
                for (let i = 0; i < lineCount; i++) {
                    rec.selectLine({ sublistId: 'item', line: i });

                    const lineItem = rec.getCurrentSublistValue({
                        sublistId: 'item', fieldId: FLD_LINE_ITEM });
                    const lineVRef = rec.getCurrentSublistValue({
                        sublistId: 'item', fieldId: FLD_LINE_VENDORREF });
                    const lineIShip = rec.getCurrentSublistText({
                        sublistId: 'item', fieldId: FLD_LINE_ISHIP });

                    log.debug('STAGE-3 compare PO line ' + i,
                        'item=' + lineItem + '/' + p.itemId +
                        ' vRef=' + lineVRef + '/' + p.vendorRef +
                        ' iship=' + lineIShip + '/' + p.shipment);

                    if (String(lineItem)  === String(p.itemId)    &&
                        String(lineVRef)  === String(p.vendorRef) &&
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

                // ---- open inventory detail subrecord on PO line ----
                rec.selectLine({ sublistId: 'item', line: idx });
                const inv = rec.getCurrentSublistSubrecord({
                    sublistId: 'item', fieldId: 'inventorydetail'
                });
                log.debug('STAGE-3 inv detail subrecord opened on PO line', idx);

                // ---- add lot assignments ----
                p.lots.forEach((lt, j) => {
                    log.debug('STAGE-3 adding lot ' + j,
                        lt.lot + ' qty=' + lt.qty + ' exp=' + lt.exp);

                    if (j === 0) {
                        inv.selectLine({ sublistId: 'inventoryassignment', line: 0 });
                    } else {
                        inv.selectNewLine({ sublistId: 'inventoryassignment' });
                    }

                    inv.setCurrentSublistValue({
                        sublistId: 'inventoryassignment',
                        fieldId  : 'receiptinventorynumber',
                        value    : lt.lot
                    });
                    inv.setCurrentSublistValue({
                        sublistId: 'inventoryassignment',
                        fieldId  : 'quantity',
                        value    : lt.qty
                    });

                    if (lt.exp) {
                        const d = parseDDMMYYYY(lt.exp);
                        if (d) {
                            inv.setCurrentSublistValue({
                                sublistId: 'inventoryassignment',
                                fieldId  : 'expirationdate',
                                value    : d
                            });
                            log.debug('  exp set', d.toString());
                        } else {
                            log.error('  bad exp', lt.exp);
                        }
                    }

                    inv.commitLine({ sublistId: 'inventoryassignment' });
                    log.debug('  committed assignment line');
                });

                rec.commitLine({ sublistId: 'item' });
                log.audit('STAGE-3 PO item line committed', 'line=' + idx);
            });

            log.debug('STAGE-3 about to save PO', poNumber);
            const savedId = rec.save({ ignoreMandatoryFields: true });
            log.audit('STAGE-3 ◀◀◀ PO SAVED', 'po=' + poNumber + ' id=' + savedId);

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
            log.error('STAGE-4 MAP ERR key=' + k, e); return true;
        });
        s.reduceSummary.errors.iterator().each((k, e) => {
            log.error('STAGE-4 REDUCE ERR key=' + k, e); return true;
        });

        log.audit('STAGE-4 ◀◀◀', 'summarize END');
    };

    // =========================================================
    // helpers
    // =========================================================
    const parseCsv = (line) => {
        const out = []; let cur = '', q = false;
        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '"') { q = !q; continue; }
            if (ch === ',' && !q) { out.push(cur); cur = ''; continue; }
            cur += ch;
        }
        out.push(cur);
        return out;
    };

    const toNum = (s) => {
        const n = parseFloat((s || '').replace(/,/g, ''));
        return isNaN(n) ? 0 : n;
    };

    const parseDDMMYYYY = (s) => {
        const m = (s || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (!m) return null;
        return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
    };

    return { getInputData, map, reduce, summarize };
});