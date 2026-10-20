/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 * Adds "Show Uncommitted" button on Sales Order (View mode).
 * Popup lists uncommitted inventory lines. Each line expands to show:
 *   - Stock by location
 *   - Other open orders holding this item (committed / allocated)
 *   - Incoming purchase orders
 */
define(['N/ui/serverWidget', 'N/search', 'N/query', 'N/url', 'N/log'], (ui, search, query, url, log) => {

    // NetSuite-generated link to an order reservation (falls back to generic transaction URL)
    const resvUrl = (id) => {
        try { return url.resolveRecord({ recordType: 'orderreservation', recordId: id }); }
        catch (e) { return '/app/accounting/transactions/transaction.nl?id=' + id; }
    };

    const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, c =>
        ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    const n = (v) => Number(v) || 0;
    const link = (id, text) => `<a href="/app/accounting/transactions/transaction.nl?id=${id}" target="_blank">${esc(text)}</a>`;

    // Runs a search; returns null instead of breaking the page if a field isn't valid in this account
    const runSearch = (type, filters, columns, mapFn) => {
        try {
            return search.create({ type, filters, columns })
                .run().getRange({ start: 0, end: 1000 }).map(mapFn);
        } catch (e) {
            log.error('Search failed: ' + type, e.message);
            return null;
        }
    };

    const beforeLoad = (ctx) => {
        if (ctx.type !== ctx.UserEventType.VIEW) return;

        const so = ctx.newRecord;
        const form = ctx.form;
        const headerLoc = so.getValue('location');
        const lines = [];

        // ---------- 1. Uncommitted lines on this SO ----------
        const count = so.getLineCount({ sublistId: 'item' });
        for (let i = 0; i < count; i++) {
            const val = (f) => so.getSublistValue({ sublistId: 'item', fieldId: f, line: i });
            const type = val('itemtype');
            if (type !== 'InvtPart' && type !== 'Assembly') continue;

            const qty = n(val('quantity'));
            const committed = n(val('quantitycommitted'));
            const fulfilled = n(val('quantityfulfilled'));
            const uncommitted = qty - committed - fulfilled;
            if (uncommitted <= 0) continue;

            lines.push({
                line: i + 1,
                itemId: String(val('item')),
                item: so.getSublistText({ sublistId: 'item', fieldId: 'item', line: i }),
                locId: String(val('location') || headerLoc || ''),
                qty, committed, fulfilled, uncommitted,
                allocated: n(val('quantitydemandallocated'))
            });
        }

        const itemIds = [...new Set(lines.map(l => l.itemId))];
        const locStock = {};   // itemId -> [location rows]
        const totals = {};     // itemId -> totals
        const orders = {};     // itemId -> [open orders holding stock]
        const pos = {};        // itemId -> [incoming POs]
        const resv = {};       // itemId -> [order reservations]
        const push = (map, k, v) => (map[k] = map[k] || []).push(v);

        if (itemIds.length) {

            // ---------- 2. Stock by location ----------
            runSearch('item',
                [['internalid', 'anyof', itemIds]],
                ['inventorylocation', 'locationquantityonhand', 'locationquantityavailable',
                 'locationquantitycommitted', 'locationquantityonorder', 'locationquantitybackordered',
                 'locationquantityintransit', 'quantityonhand', 'quantityavailable',
                 'quantitycommitted', 'quantityonorder', 'quantitybackordered'],
                r => {
                    const row = {
                        locId: String(r.getValue('inventorylocation')),
                        loc: r.getText('inventorylocation'),
                        onHand: n(r.getValue('locationquantityonhand')),
                        available: n(r.getValue('locationquantityavailable')),
                        committed: n(r.getValue('locationquantitycommitted')),
                        onOrder: n(r.getValue('locationquantityonorder')),
                        backOrdered: n(r.getValue('locationquantitybackordered')),
                        inTransit: n(r.getValue('locationquantityintransit'))
                    };
                    push(locStock, r.id, row);
                    totals[r.id] = {
                        loc: 'All Locations',
                        onHand: n(r.getValue('quantityonhand')),
                        available: n(r.getValue('quantityavailable')),
                        committed: n(r.getValue('quantitycommitted')),
                        onOrder: n(r.getValue('quantityonorder')),
                        backOrdered: n(r.getValue('quantitybackordered')),
                        inTransit: 0
                    };
                });

            // ---------- 3. Open orders holding this item (committed / allocated) ----------
            const orderFilters = (formula) => [
                ['type', 'anyof', 'SalesOrd', 'TrnfrOrd'], 'AND',
                ['mainline', 'is', 'F'], 'AND',
                ['closed', 'is', 'F'], 'AND',
                ['item', 'anyof', itemIds], 'AND',
                ['formulanumeric: ' + formula, 'greaterthan', '0']
            ];
            const orderCols = ['tranid', 'trandate', 'type', 'mainname', 'entity', 'item', 'location',
                               'quantity', 'quantitycommitted', 'quantityshiprecv'];
            const mapOrder = (withAlloc) => r => ({
                itemId: String(r.getValue('item')),
                id: r.id,
                tranid: r.getValue('tranid'),
                type: r.getText('type'),
                date: r.getValue('trandate'),
                entity: r.getText('mainname') || r.getValue('mainname') || r.getText('entity'),
                loc: r.getText('location'),
                qty: Math.abs(n(r.getValue('quantity'))),
                committed: Math.abs(n(r.getValue('quantitycommitted'))),
                allocated: withAlloc ? Math.abs(n(r.getValue('quantityallocated'))) : '-',
                shipped: Math.abs(n(r.getValue('quantityshiprecv')))
            });

            // Try with Allocated (Supply Allocation accounts); fall back to Committed only
            let orderRows = runSearch('transaction',
                orderFilters('ABS(NVL({quantitycommitted},0)) + ABS(NVL({quantityallocated},0))'),
                orderCols.concat('quantityallocated'), mapOrder(true));
            if (orderRows === null) {
                orderRows = runSearch('transaction',
                    orderFilters('ABS(NVL({quantitycommitted},0))'),
                    orderCols, mapOrder(false)) || [];
            }
            orderRows.forEach(o => push(orders, o.itemId, o));

            // ---------- 4. Incoming purchase orders ----------
            (runSearch('purchaseorder',
                [['mainline', 'is', 'F'], 'AND',
                 ['item', 'anyof', itemIds], 'AND',
                 ['status', 'anyof', 'PurchOrd:B', 'PurchOrd:D', 'PurchOrd:E']],
                ['tranid', 'trandate', 'mainname', 'entity', 'item', 'location',
                 'quantity', 'quantityshiprecv', 'expectedreceiptdate'],
                r => ({
                    itemId: String(r.getValue('item')),
                    id: r.id,
                    tranid: r.getValue('tranid'),
                    date: r.getValue('trandate'),
                    vendor: r.getText('mainname') || r.getValue('mainname') || r.getText('entity'),
                    loc: r.getText('location'),
                    qty: n(r.getValue('quantity')),
                    received: n(r.getValue('quantityshiprecv')),
                    expected: r.getValue('expectedreceiptdate')
                })) || [])
                .filter(p => p.qty - p.received > 0)
                .forEach(p => push(pos, p.itemId, p));

            // ---------- 5. Open order reservations ----------
            try {
                const ids = itemIds.map(Number).filter(Boolean).join(',');
                query.runSuiteQL({
                    query: `SELECT r.id, r.transactionnumber, r.name, r.item, r.location,
                                   BUILTIN.DF(r.location) AS locname,
                                   BUILTIN.DF(r.saleschannel) AS channel,
                                   r.startdate, r.enddate, r.targetquantity, r.quantity,
                                   r.allocatedsupply, r.quantitycommitted
                            FROM orderreservation r
                            WHERE r.item IN (${ids}) AND NVL(r.closed, 'F') = 'F'
                            ORDER BY r.startdate`
                }).asMappedResults().forEach(r => push(resv, String(r.item), {
                    id: r.id,
                    num: r.transactionnumber || r.id,
                    name: r.name,
                    channel: r.channel,
                    locId: String(r.location || ''),
                    loc: r.locname,
                    start: r.startdate,
                    end: r.enddate,
                    target: n(r.targetquantity),
                    reserved: n(r.quantity),
                    allocated: n(r.allocatedsupply),
                    committed: n(r.quantitycommitted)
                }));
            } catch (e) {
                log.error('Order reservation query failed', e.message);
            }
        }

        // ---------- 6. Build popup HTML ----------
        const COLS = 16;
        const table = (heads, body, empty) =>
            `<table class="uc_sub"><tr>${heads.map(h => `<th>${h}</th>`).join('')}</tr>` +
            (body || `<tr><td colspan="${heads.length}" class="uc_empty">${empty}</td></tr>`) + '</table>';

        let rows = lines.map(l => {
            const locs = locStock[l.itemId] || [];
            const d = locs.find(x => x.locId === l.locId) || totals[l.itemId] || {};

            // Stock by location (skip all-zero rows)
            const locBody = locs
                .filter(x => x.onHand || x.available || x.committed || x.onOrder || x.backOrdered || x.inTransit)
                .map(x => `<tr${x.locId === l.locId ? ' class="uc_hl"' : ''}>
                    <td class="l">${esc(x.loc)}</td><td>${x.onHand}</td><td>${x.available}</td>
                    <td>${x.committed}</td><td>${x.onOrder}</td><td>${x.backOrdered}</td><td>${x.inTransit}</td></tr>`)
                .join('');

            // Open orders holding stock
            const ordBody = (orders[l.itemId] || []).map(o => `<tr${String(o.id) === String(so.id) ? ' class="uc_hl"' : ''}>
                    <td class="l">${link(o.id, o.tranid)}</td><td class="l">${esc(o.type)}</td><td class="l">${esc(o.date)}</td>
                    <td class="l">${esc(o.entity)}</td><td class="l">${esc(o.loc)}</td>
                    <td>${o.qty}</td><td>${o.committed}</td><td>${o.allocated}</td><td>${o.shipped}</td></tr>`)
                .join('');

            // Order reservations
            const resList = resv[l.itemId] || [];
            const reservedHere = resList
                .filter(r => !l.locId || r.locId === l.locId)
                .reduce((s, r) => s + r.reserved, 0);
            const resBody = resList.map(r => `<tr${r.locId === l.locId ? ' class="uc_hl"' : ''}>
                    <td class="c"><a href="${resvUrl(r.id)}" target="_blank">${esc(r.id)}</a></td>
                    <td class="l"><a href="${resvUrl(r.id)}" target="_blank">${esc(r.name || r.num)}</a></td>
                    <td class="l">${esc(r.channel)}</td><td class="l">${esc(r.loc)}</td>
                    <td class="l">${esc(r.start)}</td><td class="l">${esc(r.end)}</td>
                    <td>${r.target}</td><td>${r.reserved}</td><td>${r.allocated}</td><td>${r.committed}</td></tr>`)
                .join('');

            // Incoming POs
            const poBody = (pos[l.itemId] || []).map(p => `<tr>
                    <td class="l">${link(p.id, p.tranid)}</td><td class="l">${esc(p.date)}</td>
                    <td class="l">${esc(p.vendor)}</td><td class="l">${esc(p.loc)}</td>
                    <td>${p.qty}</td><td>${p.received}</td><td>${p.qty - p.received}</td>
                    <td class="l">${esc(p.expected)}</td></tr>`)
                .join('');

            return `<tr class="uc_main">
                <td class="c"><span class="uc_tg" onclick="ucToggle(this)">&#9654;</span></td>
                <td>${l.line}</td><td class="l">${esc(l.item)}</td><td class="l">${esc(d.loc)}</td>
                <td>${l.qty}</td><td>${l.allocated}</td><td>${l.committed}</td><td>${l.fulfilled}</td>
                <td class="uc_red">${l.uncommitted}</td>
                <td>${n(d.onHand)}</td><td>${n(d.available)}</td><td>${n(d.committed)}</td>
                <td>${reservedHere}</td><td>${n(d.onOrder)}</td><td>${n(d.backOrdered)}</td><td>${n(d.inTransit)}</td>
            </tr>
            <tr class="uc_detail" style="display:none"><td colspan="${COLS}">
                <h4>Stock by Location</h4>
                ${table(['Location', 'On Hand', 'Available', 'Committed', 'On Order', 'Back Ordered', 'In Transit'],
                        locBody, 'No stock at any location')}
                <h4>Open Orders Holding This Item</h4>
                ${table(['Order #', 'Type', 'Date', 'Customer', 'Location', 'Qty', 'Committed', 'Allocated', 'Fulfilled'],
                        ordBody, 'No other orders have this item committed or allocated')}
                <h4>Order Reservations</h4>
                ${table(['Internal ID', 'Reservation', 'Sales Channel', 'Location', 'Start', 'End', 'Target Qty', 'Reserved Qty', 'Allocated Supply', 'Committed'],
                        resBody, 'No open order reservations for this item')}
                <h4>Incoming Purchase Orders</h4>
                ${table(['PO #', 'Date', 'Vendor', 'Location', 'Qty', 'Received', 'Remaining', 'Expected Receipt'],
                        poBody, 'No open purchase orders')}
            </td></tr>`;
        }).join('');

        if (!rows) rows = `<tr><td colspan="${COLS}" class="uc_empty">All items on this order are committed.</td></tr>`;

        const html = `
<style>
 #uc_overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999}
 #uc_box{background:#fff;margin:4% auto;width:92%;max-height:84%;overflow:auto;border-radius:6px;padding:16px;font-family:Arial;font-size:12px}
 #uc_box table{border-collapse:collapse;width:100%}
 #uc_box th,#uc_box td{border:1px solid #ccc;padding:5px 6px;text-align:right}
 #uc_box th{background:#607799;color:#fff;text-align:center}
 #uc_box td.l{text-align:left} #uc_box td.c{text-align:center}
 #uc_box .uc_red{color:#c00;font-weight:bold}
 #uc_box .uc_tg{cursor:pointer;color:#607799;font-size:13px;user-select:none}
 #uc_box .uc_detail>td{background:#f6f8fb;padding:10px 14px;text-align:left}
 #uc_box .uc_sub{margin-bottom:10px;background:#fff}
 #uc_box .uc_sub th{background:#9aaac2}
 #uc_box .uc_hl td{background:#fff6d6;font-weight:bold}
 #uc_box .uc_empty{text-align:center;color:#888}
 #uc_box h4{margin:8px 0 4px;color:#333}
 #uc_close{float:right;cursor:pointer;font-size:20px;font-weight:bold}
 #uc_all{margin-left:12px;cursor:pointer;font-size:12px}
</style>
<div id="uc_overlay" onclick="if(event.target===this)this.style.display='none'">
 <div id="uc_box">
  <span id="uc_close" onclick="document.getElementById('uc_overlay').style.display='none'">&times;</span>
  <h3 style="margin-top:0">Uncommitted Items <button id="uc_all" type="button" onclick="ucToggleAll(this)">Expand All</button></h3>
  <table>
   <tr>
    <th></th><th>Line</th><th>Item</th><th>Location</th>
    <th>SO Qty</th><th>Allocated</th><th>Committed</th><th>Fulfilled</th><th>Uncommitted</th>
    <th>On Hand</th><th>Available</th><th>Committed (All Orders)</th>
    <th>Reserved</th><th>On Order</th><th>Back Ordered</th><th>In Transit</th>
   </tr>
   ${rows}
  </table>
  <p style="color:#888;margin:8px 0 0">Highlighted = this order / this order's location.</p>
 </div>
</div>
<script>
 window.showUncommitted = function(){ document.getElementById('uc_overlay').style.display='block'; };
 window.ucToggle = function(el, force){
   var row = el.closest('tr').nextElementSibling;
   var open = typeof force === 'boolean' ? force : row.style.display === 'none';
   row.style.display = open ? 'table-row' : 'none';
   el.innerHTML = open ? '&#9660;' : '&#9654;';
 };
 window.ucToggleAll = function(btn){
   var open = btn.innerHTML === 'Expand All';
   document.querySelectorAll('#uc_box .uc_tg').forEach(function(t){ ucToggle(t, open); });
   btn.innerHTML = open ? 'Collapse All' : 'Expand All';
 };
</script>`;

        form.addField({
            id: 'custpage_uncommitted_html',
            type: ui.FieldType.INLINEHTML,
            label: ' '
        }).defaultValue = html;

        form.addButton({
            id: 'custpage_show_uncommitted',
            label: 'Show Uncommitted',
            functionName: 'showUncommitted'
        });
    };

    return { beforeLoad };
});