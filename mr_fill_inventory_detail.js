/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 *
 * Fill Inventory Detail on Inbound Shipment from CSV.
 * Heavy logging version — trace every step.
 */
define(['N/file', 'N/record', 'N/search'],
(file, record, search) => {

    const FILE_ID = 1278609;

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

        // Parse
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

        // Group
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
                groups[key].lots.push({ lot: r.lot, exp: r.exp, qty: r.qtyExp });
                log.debug('  + LOT', r.lot + ' qty=' + r.qtyExp);
            } else {
                log.debug('  row has no PO and no Lot — ignored', JSON.stringify(r));
            }
        });

        const list = Object.values(groups);
        log.audit('STAGE-1 groups built', list.length);

        // Log every group so we know exactly what goes into map
        list.forEach((g, i) => {
            log.audit('STAGE-1 group[' + i + ']', JSON.stringify(g));
        });

        log.audit('STAGE-1 ◀◀◀', 'getInputData END, returning ' + list.length);
        return list;
    };

    // =========================================================
    // 2) map
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

            // Greedy split
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
            }

            // Emit
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
                const keyOut = g.shipment || 'UNKNOWN';
                log.audit('STAGE-2 EMIT', 'key=' + keyOut + ' value=' + JSON.stringify(payload));
                ctx.write({ key: keyOut, value: JSON.stringify(payload) });
                emitted++;
            });

            log.audit('STAGE-2 ◀◀◀', 'map END item=' + g.itemId + ' emitted=' + emitted);

        } catch (e) {
            log.error('STAGE-2 map EXCEPTION', e.message + ' stack=' + e.stack);
        }
    };

    // =========================================================
    // 3) reduce
    // =========================================================
    const reduce = (ctx) => {
        log.audit('STAGE-3 ▶▶▶', 'reduce START key=' + ctx.key +
            ' values=' + (ctx.values ? ctx.values.length : 'NONE'));

        try {
            const shipmentNum = ctx.key;

            // Log every payload we received
            ctx.values.forEach((v, i) => {
                log.debug('STAGE-3 incoming[' + i + ']', v);
            });

            // Find inbound shipment
            log.debug('STAGE-3 searching inbound shipment', shipmentNum);
            const hits = search.create({
                type: 'inboundshipment',
                filters: [['shipmentnumber', 'is', shipmentNum]],
                columns: ['internalid']
            }).run().getRange({ start: 0, end: 5 });

            log.audit('STAGE-3 search result', 'hits=' + hits.length);

            if (!hits.length) {
                log.error('STAGE-3 NOT FOUND', 'shipment=' + shipmentNum);
                return;
            }

            const shipId = hits[0].getValue('internalid');
            log.audit('STAGE-3 shipment id', shipId);

            const rec = record.load({
                type: 'inboundshipment', id: shipId, isDynamic: true
            });
            const lineCount = rec.getLineCount({ sublistId: 'items' });
            log.audit('STAGE-3 line count', lineCount);

            // Log every existing line so we can see what we're matching against
            for (let i = 0; i < lineCount; i++) {
                rec.selectLine({ sublistId: 'items', line: i });
                const it = rec.getCurrentSublistValue({ sublistId: 'items', fieldId: 'item' });
                const vr = rec.getCurrentSublistValue({ sublistId: 'items', fieldId: 'custcol_mi_vendor_ref_number' });
                const pt = rec.getCurrentSublistText({ sublistId: 'items', fieldId: 'purchaseorder' });
                log.debug('STAGE-3 existing line ' + i,
                    'item=' + it + ' vRef=' + vr + ' PO=' + pt);
            }

            // Process each payload
            ctx.values.forEach((v, pIdx) => {
                log.audit('STAGE-3 processing payload ' + pIdx, v);
                const p = JSON.parse(v);

                let idx = -1;
                for (let i = 0; i < lineCount; i++) {
                    rec.selectLine({ sublistId: 'items', line: i });
                    const lineItem = rec.getCurrentSublistValue({
                        sublistId: 'items', fieldId: 'itemid' });
                    const lineVRef = rec.getCurrentSublistValue({
                        sublistId: 'items', fieldId: 'custrecord_mi_vendor_ref_number' });
                    let linePO = rec.getCurrentSublistText({
                        sublistId: 'items', fieldId: 'purchaseorder' });
                    if (linePO && linePO.indexOf('PO#') != -1) {
                        linePO = linePO.replace('PO#', '');
                    }

                    log.debug('STAGE-3 compare line ' + i,
                        'li=' + lineItem + '/' + p.itemId +
                        ' vr=' + lineVRef + '/' + p.vendorRef +
                        ' po=' + linePO + '/' + p.po);

                    if (String(lineItem) === String(p.itemId) &&
                        String(lineVRef) === String(p.vendorRef) &&
                        String(linePO)   === String(p.po)) {
                        idx = i;
                        log.audit('STAGE-3 MATCHED', 'line=' + i);
                        break;
                    }
                }

                if (idx === -1) {
                    log.error('STAGE-3 NO LINE MATCH',
                        'item=' + p.itemId + ' PO=' + p.po + ' vRef=' + p.vendorRef);
                    return;
                }

                // Open inventory detail
                rec.selectLine({ sublistId: 'items', line: idx });
                const inv = rec.getCurrentSublistSubrecord({
                    sublistId: 'items', fieldId: 'inventorydetail'
                });
                log.debug('STAGE-3 inv detail subrecord opened for line', idx);

                // Add lots
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

                rec.commitLine({ sublistId: 'items' });
                log.audit('STAGE-3 item line committed', 'line=' + idx);
            });

            log.debug('STAGE-3 about to save', 'shipment=' + shipmentNum);
            const savedId = rec.save({ ignoreMandatoryFields: true });
            log.audit('STAGE-3 ◀◀◀ SAVED', 'shipment=' + shipmentNum + ' id=' + savedId);

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
        log.audit('STAGE-4 counts',
            'mapKeys=' + s.mapSummary.keys.iterator + // just a reference
            ' mapErrors=' + s.mapSummary.errors.iterator +
            ' reduceErrors=' + s.reduceSummary.errors.iterator);

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