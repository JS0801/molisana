/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 *
 * Fill Inventory Detail on Inbound Shipment from CSV.
 *
 * Match the PO line using:
 *   Column1        -> custcol_mi_vendor_ref_number
 *   Item ID        -> item
 *   Purchase Order -> purchaseorder   (e.g. PO3549)
 *
 * A lot qty can span two POs of the same item, so we walk lots
 * in order and fill each PO's qty until it's complete.
 */
define(['N/file', 'N/record', 'N/search', 'N/runtime'],
(file, record, search, runtime) => {


    // =========================================================
    // 1) getInputData  — read CSV, group by item per shipment
    // =========================================================
    const getInputData = () => {
        log.audit('STAGE 1', 'getInputData START');

        const lines = file.load({ id: 1278609 })
            .getContents().split(/\r?\n/);
        log.debug('CSV line count', lines.length);

        // Parse rows (skip header on line 0)
        const rows = [];
        for (let i = 1; i < lines.length; i++) {
            if (!lines[i].trim()) continue;
            const c = parseCsv(lines[i]);
            rows.push({
                shipment : c[0],
                vendorRef: c[1],             // Column1
                po       : c[5],             // Purchase Order
                itemId   : c[6],
                lot      : c[13],
                exp      : c[14],
                qtyExp   : toNum(c[15]),
                qtyRec   : toNum(c[16])
            });
        }
        log.debug('Parsed rows', rows.length);

        // Carry down shipment + vendorRef (blank on lot rows)
        let curShip = '', curVRef = '';
        const groups = {};

        rows.forEach(r => {
            if (r.shipment)  curShip = r.shipment;
            if (r.vendorRef) curVRef = r.vendorRef;

            const key = curShip + '|' + curVRef + '|' + r.itemId;
            if (!groups[key]) {
                groups[key] = {
                    shipment : curShip,
                    vendorRef: curVRef,
                    itemId   : r.itemId,
                    pos      : [],   // [{po, qty}]
                    lots     : []    // [{lot, exp, qty}]
                };
            }

            if (r.po) {
                groups[key].pos.push({ po: r.po, qty: r.qtyExp });
            } else if (r.lot) {
                groups[key].lots.push({
                    lot: r.lot, exp: r.exp, qty: r.qtyRec
                });
            }
        });

        const list = Object.values(groups);
        log.audit('STAGE 1', 'Groups = ' + list.length);
        return list;
    };

    // =========================================================
    // 2) map  — split lots across POs, emit per PO allocation
    // =========================================================
    const map = (ctx) => {
        const g = JSON.parse(ctx.value);
        log.audit('STAGE 2 map',
            'ship=' + g.shipment + ' item=' + g.itemId +
            ' POs=' + g.pos.length + ' Lots=' + g.lots.length);


      try {
        
        // Working copies with "remaining" counters
        const pos  = g.pos.map(p  => ({ po: p.po, remaining: p.qty, lots: [] }));
        const lots = g.lots.map(l => ({ lot: l.lot, exp: l.exp, remaining: l.qty }));

        // Walk lots, fill each PO's qty in order
        let pi = 0;
        for (const lot of lots) {
            while (lot.remaining > 0 && pi < pos.length) {
                if (pos[pi].remaining <= 0) { pi++; continue; }

                const take = Math.min(lot.remaining, pos[pi].remaining);
                pos[pi].lots.push({ lot: lot.lot, exp: lot.exp, qty: take });
                pos[pi].remaining -= take;
                lot.remaining     -= take;

                log.debug('Split', 'item=' + g.itemId +
                    ' ' + take + ' of ' + lot.lot + ' → ' + pos[pi].po);

                if (pos[pi].remaining <= 0) pi++;
            }
        }

        // Emit one payload per PO, keyed by shipment
        pos.forEach(p => {
            if (!p.lots.length) return;
            const payload = {
                shipment : g.shipment,
                vendorRef: g.vendorRef,
                itemId   : g.itemId,
                po       : p.po,
                lots     : p.lots
            };
            log.debug('Emit', JSON.stringify(payload));
            ctx.write({ key: g.shipment, value: JSON.stringify(payload) });
        });

      } catch (error) {
        log.error('Map Error', error)
      }
    };

    // =========================================================
    // 3) reduce  — for each shipment, fill inventory detail
    // =========================================================
    const reduce = (ctx) => {
        const shipmentNum = ctx.key;
        log.audit('STAGE 3 reduce',
            'shipment=' + shipmentNum + ' payloads=' + ctx.values.length);

      try {
                // Find the inbound shipment
        const hits = search.create({
            type: 'inboundshipment',
            filters: [['shipmentnumber', 'is', shipmentNum]],
            columns: ['internalid']
        }).run().getRange({ start: 0, end: 1 });

        if (!hits.length) {
            log.error('Shipment not found', shipmentNum);
            return;
        }

        const shipId = hits[0].getValue('internalid');
        log.debug('Shipment ID', shipId);

        const rec = record.load({
            type: 'inboundshipment', id: shipId, isDynamic: true
        });
        const lineCount = rec.getLineCount({ sublistId: 'items' });
        log.debug('Line count', lineCount);

        ctx.values.forEach(v => {
            const p = JSON.parse(v);
            log.audit('Filling',
                'item=' + p.itemId + ' PO=' + p.po + ' vRef=' + p.vendorRef);

            // ---- locate the matching line ----
            let idx = -1;
            for (let i = 0; i < lineCount; i++) {
                rec.selectLine({ sublistId: 'items', line: i });

                const lineItem = rec.getCurrentSublistValue({
                    sublistId: 'items', fieldId: 'item' });
                const lineVRef = rec.getCurrentSublistValue({
                    sublistId: 'items', fieldId: 'custcol_mi_vendor_ref_number' });
                const linePO = rec.getCurrentSublistText({
                    sublistId: 'items', fieldId: 'purchaseorder' });

                if (String(lineItem) === String(p.itemId) &&
                    String(lineVRef) === String(p.vendorRef) &&
                    String(linePO)   === String(p.po)) {
                    idx = i;
                    break;
                }
            }

            if (idx === -1) {
                log.error('Line not matched',
                    'item=' + p.itemId + ' PO=' + p.po);
                return;
            }
            log.debug('Matched line', idx);

            // ---- open inventory detail subrecord ----
            rec.selectLine({ sublistId: 'items', line: idx });
            const inv = rec.getCurrentSublistSubrecord({
                sublistId: 'items', fieldId: 'inventorydetail'
            });

            // ---- add each lot ----
            p.lots.forEach((lt, j) => {
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
                    }
                }

                inv.commitLine({ sublistId: 'inventoryassignment' });
                log.debug('  lot added', lt.lot + ' qty=' + lt.qty);
            });

            rec.commitLine({ sublistId: 'items' });
        });

    //    const savedId = rec.save({ ignoreMandatoryFields: true });
        log.audit('SAVED', 'shipment=' + shipmentNum + ' id=' + savedId);
      } catch (error) {
        log.error('Reduce Error', error)
      }


    };

    // =========================================================
    // 4) summarize
    // =========================================================
    const summarize = (s) => {
        log.audit('STAGE 4 summarize',
            'map=' + s.mapSummary.usage +
            ' reduce=' + s.reduceSummary.usage);

        s.mapSummary.errors.iterator().each((k, e) => {
            log.error('map error ' + k, e); return true;
        });
        s.reduceSummary.errors.iterator().each((k, e) => {
            log.error('reduce error ' + k, e); return true;
        });
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

    // "18/02/2027" -> Date
    const parseDDMMYYYY = (s) => {
        const m = (s || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (!m) return null;
        return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
    };

    return { getInputData, map, reduce, summarize };
});
