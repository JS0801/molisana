/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 *
 */
define(['N/search', 'N/file', 'N/log', 'N/runtime'],
    (search, file, log, runtime) => {
        'use strict';

        const CONFIG = {
            MASTER_FOLDER_ID: 493530,
            MASTER_DATE_FIELD: 'created',
            MAX_FILE_BYTES: 9 * 1024 * 1024,
            MAX_MASTER_ROWS: 10000,
            MAX_DISPLAY_ROWS: 15000,
            SALES_ORDER_LOOKBACK_DAYS: 35
        };

        const MASTER_MAP = [
            ['LCL - XMAS', 'TTL LCL Booked', 'TTL LCL Filled', 'Total LCL Balance'],
            ['Amazon - XMAS', 'Amazon Booked', 'Amazon Filled', 'Amazon Balance'],
            ['Dollar Tree - XMAS', 'DT Booked', 'DT Filled', 'DT Balance'],
            ['Rexall - XMAS', 'RX Booked', 'RX Filled', 'RX Balance'],
            ['Metro - XMAS', 'Metro Committed', 'Metro Sold', 'Metro Remaining'],
            ['Food Basics - XMAS', 'Basics Committed', 'Basics Sold', 'Basics Remaining'],
            ['Adonis - XMAS', 'Adonis Booked', 'Adonis Filled', 'Adonis Balance'],
            ['Giant Tiger - XMAS', 'GT Booked', 'GT Filled', 'GT Balance'],
            ['Sobeys/ Fresco - XMAS', 'Sobeys Booked', 'Sobeys Filled', 'Sobeys Balance'],
            ['Costco - XMAS', 'Costco Booked', 'Costco Filled', 'Costco Balance'],
            ["Longo's - XMAS", 'Longos Booked', 'Longos Filled', 'Longos Balance'],
            ['Walmart - XMAS', 'Walmart Booked', 'Walmart Filled', 'Walmart Balance'],
            ['Eataly - XMAS', 'Eataly Booked', 'Eataly Filled', 'Eataly Balance'],
            ['Farm Boy - XMAS', 'Farm Boy Booked', 'Farm Boy Filled', 'Farm Boy Balance'],
            ['DECI - XMAS', 'Deci Booked', 'Deci Filled', 'Deci Balance'],
            ['Italian Star Deli - XMAS', 'ISD Booked', 'ISD Filled', 'ISD Balance'],
            ['MVR - XMAS', 'MVR Booked', 'MVR Filled', 'MVR Balance'],
            ['TJX - XMAS', 'TJX Baskets Booked', 'TJX Baskets Filled', 'TJX Baskets Balance'],
            ['Save On - XMAS', 'Save On Booked', 'Save On Filled', 'Save On Balance'],
            ['DGL - XMAS', 'DGL Booked', 'DGL Filled', 'DGL Balance'],
            ['Xmas 2026', 'Other Booked', 'Other Sold', 'Other Balance']
        ];

        const clean = v => String(v == null ? '' : v).replace(/\u00a0/g, ' ').trim().replace(/\s+/g, ' ');
        const key = v => clean(v).toLowerCase();
        const channelNames = Object.create(null);
        MASTER_MAP.forEach(m => { channelNames[key(m[0])] = m[0]; });
        function channelName(value) {
            const text = clean(value);
            if (!text || ['- none -', '@none@', 'none'].includes(key(text))) return 'Xmas 2026';
            return channelNames[key(text)] || 'Xmas 2026';
        }
        function escapeHtml(v) {
            return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({
                '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
            }[c]));
        }
        function numberValue(value, label, blankValue) {
            if (value == null || clean(value) === '') return blankValue;
            let text = clean(value).replace(/,/g, '');
            if (/^\(.*\)$/.test(text)) text = '-' + text.slice(1, -1);
            const number = Number(text);
            if (!Number.isFinite(number)) throw new Error('Invalid quantity in ' + label + ': ' + clean(value));
            return number;
        }
        function assertUsage() {
            if (runtime.getCurrentScript().getRemainingUsage() < 80) {
                throw new Error('This request is too large. Select a specific item and refresh. No partial totals are displayed.');
            }
        }
        function eachResult(s, callback) {
            const pages = s.runPaged({ pageSize: 1000 });
            if (pages.count > 1000000) throw new Error('Search exceeds the paging limit. Select a specific item.');
            pages.pageRanges.forEach(p => {
                assertUsage();
                pages.fetch({ index: p.index }).data.forEach(callback);
            });
        }
        function getRow(rows, id, name, channel) {
            const itemId = clean(id);
            if (!/^\d+$/.test(itemId)) throw new Error('Expected a NetSuite item internal ID, received: ' + itemId);
            const canonical = channelName(channel);
            const rowKey = itemId + '|' + key(canonical);
            if (!rows.has(rowKey)) rows.set(rowKey, {
                id: itemId, name: clean(name), description: '', channel: canonical,
                so: null, reservation: null, master: null
            });
            const row = rows.get(rowKey);
            if (!row.name && name) row.name = clean(name);
            return row;
        }
        function addQuantities(row, field, values) {
            if (!row[field]) row[field] = values.slice();
            else row[field] = row[field].map((v, i) =>
                v === null || values[i] === null ? null : v + values[i]);
        }

        function loadSalesOrders(rows, itemId) {
            const columns = [
                search.createColumn({ name: 'internalid', join: 'item', summary: 'GROUP', sort: search.Sort.ASC }),
                search.createColumn({ name: 'item', summary: 'GROUP' }),
                search.createColumn({ name: 'quantity', summary: 'SUM' }),
                search.createColumn({ name: 'quantityshiprecv', summary: 'SUM' }),
                search.createColumn({ name: 'formulanumeric', summary: 'SUM', formula: '{quantity} - NVL({quantityshiprecv},0)' }),
                search.createColumn({ name: 'custentity_default_shipping_channel', join: 'customerMain', summary: 'GROUP', sort: search.Sort.ASC }),
                search.createColumn({ name: 'category', join: 'customerMain', summary: 'GROUP', sort: search.Sort.ASC })
            ];
            const filters = [
                ['type', 'anyof', 'SalesOrd'], 'AND', ['mainline', 'is', 'F'], 'AND',
                [
                    ['custbody_order_type_cp', 'anyof', '3'], 'OR', ['memomain', 'contains', 'Xmas'], 'OR',
                    ['formulatext: {otherrefnum}', 'contains', 'CHRISTMAS'], 'OR',
                    ['formulatext: {otherrefnum}', 'contains', 'Xmas'], 'OR',
                    ['memomain', 'contains', 'CHRISTMAS'], 'OR',
                    ["formulanumeric: CASE WHEN {item} LIKE '%XMAS%' OR {item} LIKE '%BACI%' THEN 1 ELSE 0 END", 'equalto', '1'], 'OR',
                    ["formulanumeric: CASE WHEN UPPER({item.displayname}) LIKE '%CHIOSTRO%' OR UPPER({item.displayname}) LIKE '%STREGA%' THEN 1 ELSE 0 END", 'equalto', '1']
                ], 'AND', ['actualshipdate', 'within', 'daysago' + CONFIG.SALES_ORDER_LOOKBACK_DAYS, 'daysago0'], 'AND',
                ['taxline', 'is', 'F'], 'AND', ['cogs', 'is', 'F'], 'AND', ['closed', 'is', 'F']
            ];
            if (itemId) filters.push('AND', ['item', 'anyof', itemId]);
            const s = search.create({ type: 'salesorder', settings: [{ name: 'consolidationtype', value: 'ACCTTYPE' }], filters, columns });
            let withoutItem = 0;
            eachResult(s, r => {
                const id = r.getValue(columns[1]);
                if (!id || id === '- None -') { withoutItem++; return; }
                const row = getRow(rows, id, r.getText(columns[1]), r.getText(columns[5]));
                addQuantities(row, 'so', [2, 3, 4].map(i => numberValue(r.getValue(columns[i]), 'sales order result', 0)));
            });
            return withoutItem;
        }

        function loadReservations(rows, itemId) {
            const columns = [
                search.createColumn({ name: 'internalid', join: 'item', summary: 'GROUP', sort: search.Sort.ASC }),
                search.createColumn({ name: 'saleschannel', summary: 'GROUP', sort: search.Sort.ASC }),
                search.createColumn({ name: 'item', summary: 'GROUP' }),
                search.createColumn({ name: 'quantity', summary: 'SUM' }),
                search.createColumn({ name: 'quantitycommitted', summary: 'SUM' }),
                search.createColumn({ name: 'quantityshiprecv', summary: 'SUM' }),
                search.createColumn({ name: 'formulanumeric', summary: 'SUM', formula: '{quantity} - NVL({quantitycommitted},0) - NVL({quantityshiprecv},0)' })
            ];
            const filters = [
                ['type', 'anyof', 'OrdResv'], 'AND', ['closed', 'is', 'F'], 'AND', ['mainline', 'is', 'F'], 'AND',
                ['formulanumeric: {quantity} - NVL({quantitycommitted},0)', 'greaterthan', '0'], 'AND',
                ['saleschannel', 'noneof', '@NONE@', '6', '1', '2', '5', '3', '4']
            ];
            if (itemId) filters.push('AND', ['item', 'anyof', itemId]);
            const s = search.create({ type: 'orderreservation', settings: [{ name: 'consolidationtype', value: 'ACCTTYPE' }], filters, columns });
            let withoutItem = 0;
            eachResult(s, r => {
                const id = r.getValue(columns[2]);
                if (!id || id === '- None -') { withoutItem++; return; }
                const row = getRow(rows, id, r.getText(columns[2]), r.getText(columns[1]));
                addQuantities(row, 'reservation', [3, 4, 5, 6].map(i => numberValue(r.getValue(columns[i]), 'reservation result', 0)));
            });
            return withoutItem;
        }

        function newestMaster() {
            if (!Number.isInteger(Number(CONFIG.MASTER_FOLDER_ID)) || Number(CONFIG.MASTER_FOLDER_ID) <= 0) {
                throw new Error('Set CONFIG.MASTER_FOLDER_ID in mi_xmas_qty_suitelet.js to the File Cabinet folder internal ID.');
            }
            const cols = [
                search.createColumn({ name: CONFIG.MASTER_DATE_FIELD, sort: search.Sort.DESC }),
                search.createColumn({ name: 'internalid', sort: search.Sort.DESC }),
                search.createColumn({ name: 'name' })
            ];
            const found = search.create({
                type: 'file',
                filters: [
                    ['folder', 'anyof', String(CONFIG.MASTER_FOLDER_ID)]
                ], columns: cols
            }).run().getRange({ start: 0, end: 1 });
            if (!found.length) throw new Error('No .csv file found directly in master folder ' + CONFIG.MASTER_FOLDER_ID + '.');
            const f = file.load({ id: found[0].getValue(cols[1]) });
            if (f.size > CONFIG.MAX_FILE_BYTES) throw new Error('Newest master file exceeds the configured 9 MB limit: ' + f.name);
            return { file: f, date: found[0].getValue(cols[0]) };
        }

        function parseCsv(contents) {
            const text = String(contents).replace(/^\uFEFF/, '');
            if (text.includes('\u0000')) throw new Error('Save the master CSV as UTF-8, comma-delimited.');
            const rows = [];
            let row = [], value = '', quoted = false, closed = false;
            const pushCell = () => {
                row.push(value); value = ''; closed = false;
                if (row.length > 301) throw new Error('Master CSV exceeds 301 columns.');
            };
            const pushRow = () => {
                pushCell(); rows.push(row); row = [];
                if (rows.length > CONFIG.MAX_MASTER_ROWS + 1) throw new Error('Master CSV exceeds ' + CONFIG.MAX_MASTER_ROWS + ' data rows.');
            };
            for (let i = 0; i < text.length; i++) {
                const c = text[i];
                if (quoted) {
                    if (c === '"') {
                        if (text[i + 1] === '"') { value += '"'; i++; }
                        else { quoted = false; closed = true; }
                    } else value += c;
                } else if (c === ',') pushCell();
                else if (c === '\r' || c === '\n') {
                    pushRow(); if (c === '\r' && text[i + 1] === '\n') i++;
                } else if (c === '"' && !value && !closed) quoted = true;
                else {
                    if (closed || c === '"') throw new Error('Malformed CSV quoting near record ' + (rows.length + 1) + '.');
                    value += c;
                }
            }
            if (quoted) throw new Error('Master CSV contains an unclosed quoted field.');
            if (value || row.length || closed) pushRow();
            return rows;
        }

        function parseMaster(f) {
            if (!/\.csv$/i.test(f.name)) throw new Error('Master file must be a .csv file.');
            const table = parseCsv(f.getContents());
            const index = Object.create(null);
            (table[0] || []).forEach((h, i) => {
                if (!clean(h)) return;
                if (index[key(h)] !== undefined) throw new Error('Duplicate master header: ' + h);
                index[key(h)] = i;
            });
            const required = ['Item ID'].concat(...MASTER_MAP.map(m => m.slice(1, 2)));
            const missing = required.filter(h => index[key(h)] === undefined);
            if (missing.length) throw new Error('Newest master file is missing headers: ' + missing.join(', '));
            const parsed = [], seen = new Set();
            for (let r = 1; r < table.length; r++) {
                const values = table[r] || [];
                if (values.every(v => v == null || clean(v) === '')) continue;
                if (values.length !== table[0].length) throw new Error('Master CSV record ' + (r + 1) + ' has ' + values.length + ' fields; expected ' + table[0].length + '.');
                const rawId = clean(values[index[key('Item ID')]]);
                if (['total', 'grand total'].includes(key(rawId))) continue;
                if (!/^\d+$/.test(rawId) || Number(rawId) <= 0 || !Number.isSafeInteger(Number(rawId))) {
                    throw new Error('Master row ' + (r + 1) + ': Item ID must be a NetSuite item internal ID.');
                }
                const id = String(Number(rawId));
                if (seen.has(id)) throw new Error('Duplicate Item ID ' + id + ' in master row ' + (r + 1) + '. Keep one row per item.');
                seen.add(id);
                MASTER_MAP.forEach(m => {
                    const quantities = m.slice(1, 2).map(h => {
                        const col = index[key(h)];
                        return numberValue(values[col], h + ', row ' + (r + 1), null);
                    });
                    parsed.push({ id, channel: m[0], quantities });
                });
            }
            if (!seen.size) throw new Error('Master file contains no item rows.');
            return { parsed, itemCount: seen.size };
        }

        function loadMaster(rows, itemId) {
            const newest = newestMaster();
            let result;
            try { result = parseMaster(newest.file); }
            catch (e) { throw new Error(newest.file.name + ': ' + e.message); }
            result.parsed.forEach(r => {
                if (!itemId || r.id === itemId) getRow(rows, r.id, '', r.channel).master = r.quantities;
            });
            return { name: newest.file.name, id: newest.file.id, date: newest.date, itemCount: result.itemCount };
        }

        function fillItemNames(rows, warnings) {
            const ids = Array.from(new Set(Array.from(rows.values()).map(r => r.id)));
            const details = new Map();
            for (let start = 0; start < ids.length; start += 500) {
                const batch = ids.slice(start, start + 500);
                eachResult(search.create({ type: 'item', filters: [['internalid', 'anyof', batch]], columns: [
                    search.createColumn({ name: 'internalid', sort: search.Sort.ASC }),
                    'itemid', 'displayname', 'salesdescription'
                ] }), r => { details.set(String(r.id), {
                    name: clean(r.getValue({ name: 'itemid' })),
                    description: clean(r.getValue({ name: 'salesdescription' })) || clean(r.getValue({ name: 'displayname' }))
                }); });
            }
            rows.forEach(r => {
                const detail = details.get(r.id);
                if (detail) {
                    r.name = detail.name || r.name;
                    r.description = detail.description;
                }
            });
            const unresolved = ids.filter(id => !details.has(id));
            if (unresolved.length) warnings.push(unresolved.length + ' item ID(s) could not be resolved or are not visible to this role: ' + unresolved.slice(0, 20).join(', ') + '.');
        }

        // Missing or blank master bookings remain unknown, rather than becoming zero.
        // A missing sales group means no shipments matched the existing report filters.
        function commitmentSummary(row) {
            const booked = row.master && row.master[0] != null ? row.master[0] : null;
            const shipped = row.so ? row.so[1] : 0;
            if (booked === null || shipped === null) return { booked, remaining: null, status: 'Booking unavailable', css: '' };
            const difference = booked - shipped;
            const remaining = Math.abs(difference) < 0.000000001 ? 0 : difference;
            if (remaining < 0) return { booked, remaining, status: 'Over commitment', css: 'over' };
            if (booked > 0 && remaining === 0) return { booked, remaining, status: 'Completed', css: 'complete' };
            if (booked > 0 && shipped > 0) return { booked, remaining, status: 'Partially fulfilled', css: 'partial' };
            return { booked, remaining, status: booked === 0 ? 'No commitment' : 'Not shipped', css: '' };
        }

        function exportCsv(rows) {
            const cell = v => {
                if (v === null || v === undefined) return '""';
                let s = String(v);
                if (typeof v === 'string' && /^[\s]*[=+@-]/.test(s)) s = "'" + s;
                return '"' + s.replace(/"/g, '""') + '"';
            };
            const headers = ['Item ID', 'Item Description', 'Item Name', 'Sales Channel',
                'Booked / Original Commitments',
                'Sales Order Qty - Ordered', 'Sales Order Qty - Shipped', 'Sales Order Qty - Remaining',
                'Remaining Commitments', 'Commitment Status',
                'Order Reservation - Reserved', 'Order Reservation - Committed', 'Order Reservation - Used', 'Order Reservation - Remaining'];
            return '\uFEFF' + [headers].concat(rows.map(r => {
                const commitment = commitmentSummary(r);
                return [r.id, r.description, r.name, r.channel, commitment.booked].concat(
                    r.so || [0, 0, 0], [commitment.remaining, commitment.status], r.reservation || [null, null, null, null]
                );
            })).map(r => r.map(cell).join(',')).join('\r\n');
        }

        function renderTable(rows) {
            const fmt = n => n === null ? '—' : escapeHtml(n.toLocaleString('en-US', { maximumFractionDigits: 6 }));
            const group = (v, count) => (v || Array(count).fill(null)).map(n =>
                '<td class="num' + (n !== null && n < 0 ? ' negative' : '') + '">' + fmt(n) + '</td>').join('');
            return '<div class="table-wrap"><table><thead><tr>' +
                '<th colspan="4">Item Details</th><th class="master">Original Booking</th><th colspan="3" class="so">Sales Order Qty</th><th colspan="2" class="master">Commitment Progress</th><th colspan="4" class="reservation">Order Reservation</th>' +
                '</tr><tr><th>Item ID</th><th>Item Description</th><th>Item Name</th><th>Sales Channel</th><th>Booked / Original Commitments</th><th>Ordered</th><th>Shipped</th><th>Remaining</th><th>Remaining Commitments</th><th>Status</th><th>Reserved</th><th>Committed</th><th>Used</th><th>Remaining</th></tr></thead><tbody>' +
                (rows.map(r => {
                    const commitment = commitmentSummary(r);
                    return '<tr class="' + commitment.css + '"><td>' + escapeHtml(r.id) + '</td><td class="description">' + escapeHtml(r.description || '—') + '</td><td>' + escapeHtml(r.name || 'Unresolved item') + '</td><td>' + escapeHtml(r.channel) + '</td>' +
                        group([commitment.booked], 1) + group(r.so || [0, 0, 0], 3) + group([commitment.remaining], 1) + '<td class="status">' + escapeHtml(commitment.status) + '</td>' +
                        group(r.reservation, 4) + '</tr>';
                }).join('') || '<tr><td colspan="14">No matching items.</td></tr>') + '</tbody></table></div>';
        }

        function renderPage(rows, allRows, itemId, chosenChannel, warnings, fatal) {
            const h = escapeHtml;
            const items = new Map();
            allRows.forEach(r => items.set(r.id, [r.name || r.id, r.description].filter(Boolean).join(' — ')));
            const options = Array.from(items).sort((a, b) => a[1].localeCompare(b[1])).map(([id, name]) =>
                '<option value="' + h(id) + '">' + h(name) + '</option>').join('');
            const channels = MASTER_MAP.map(m => m[0]).sort((a, b) => a.localeCompare(b)).map(c =>
                '<option value="' + h(c) + '"' + (c === chosenChannel ? ' selected' : '') + '>' + h(c) + '</option>').join('');
            return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
                <title>Christmas Quantity Comparison</title><style>
                *{box-sizing:border-box}body{margin:0;background:#f5f7fa;color:#21334b;font:14px Arial,sans-serif}main{padding:28px;max-width:100%}h1{font-size:25px;margin:0 0 20px}
                form{display:flex;align-items:flex-end;gap:14px;flex-wrap:wrap;background:white;padding:18px;border:1px solid #dce2eb;border-radius:8px}
                label{display:flex;flex-direction:column;gap:8px;font-weight:600}input,select,button{font:inherit;border:1px solid #b9c5d4;border-radius:5px;padding:10px 12px;min-height:40px}
                input,select{min-width:230px;background:white}button{cursor:pointer;background:#245885;color:white;border-color:#245885}.secondary{background:white;color:#245885}
                .note{color:#c11c2b;margin:18px 0;font-weight:600;line-height:1.6}.error{background:#fff0f0;color:#a5202d;padding:14px;margin:12px 0;border:1px solid #efbcc1}
                .table-wrap{overflow:auto;max-height:75vh;border:1px solid #d4dde7;border-radius:7px;background:white}table{border-collapse:separate;border-spacing:0;min-width:1350px;width:100%;font-size:12px}
                th,td{padding:11px 12px;border-right:1px solid #e0e6ed;border-bottom:1px solid #e0e6ed;white-space:nowrap}thead th{position:sticky;top:0;height:40px;background:#263a55;color:white;text-align:center;z-index:2}
                thead tr:nth-child(2) th{top:40px;background:#eaf0f6;color:#24364e}.so{background:#225e87}.reservation{background:#4e568c}.master{background:#247365}
                tbody tr:nth-child(even){background:#f7f9fc}tbody tr:hover{background:#edf4ff}.num{text-align:right;font-variant-numeric:tabular-nums}.negative{color:#bb2331}
                tbody tr.complete{background:#e1f3e5}tbody tr.over{background:#fde5e5}tbody tr.partial{background:#fff3cd}
                tbody tr.complete:hover{background:#cce9d3}tbody tr.over:hover{background:#f9d0d0}tbody tr.partial:hover{background:#ffe9a3}
                .status{font-weight:700}.description{white-space:normal;min-width:220px;max-width:380px}
                .legend{margin:12px 0;line-height:1.6;color:#394c63}
                @media(max-width:650px){main{padding:14px}label,input,select{width:100%}h1{font-size:21px}}
                </style></head><body><main><h1>Christmas Quantity Comparison</h1>
                <form method="post"><label>Item<input name="custpage_item" list="items" value="${h(itemId)}" placeholder="All items — enter or select item ID"><datalist id="items">${options}</datalist></label>
                <label>Sales Channel<select name="custpage_channel"><option value="">All Sales Channels</option>${channels}</select></label>
                <button type="submit" name="action" value="view">Apply Filters</button><button type="submit" name="action" value="export" class="secondary">Export CSV</button>
                <button type="submit" name="action" value="clear" class="secondary">Clear Filters</button></form>
                <p class="note">Sales Order Qty uses orders with an actual ship date in the past ${CONFIG.SALES_ORDER_LOOKBACK_DAYS} days. Remaining Commitments = Original Booked − Shipped from this report; it is not an all-season balance. Any other sales channel is merged with Xmas 2026.</p>
                <p class="legend">Green: completed · Yellow: partially fulfilled · Red: shipped over commitment. Missing bookings show as unavailable; zero bookings with shipments show as over commitment. Order Reservation quantities are shown separately.</p>
                ${warnings.map(w => '<div class="error" role="alert">' + h(w) + '</div>').join('')}
                ${fatal ? '<div class="error" role="alert">' + h(fatal) + '</div>' : renderTable(rows)}
                </main></body></html>`;
        }

        function onRequest(context) {
            const params = context.request.parameters || {};
            const itemId = params.action === 'clear' ? '' : clean(params.custpage_item);
            const channelInput = params.action === 'clear' ? '' : clean(params.custpage_channel);
            const chosenChannel = channelInput ? channelName(channelInput) : '';
            const rows = new Map(), warnings = [];
            let fatal = null;
            try {
                if (itemId && !/^\d+$/.test(itemId)) throw new Error('Select or enter a valid item internal ID.');
                const omitted = loadSalesOrders(rows, itemId) + loadReservations(rows, itemId);
                if (omitted) warnings.push(omitted + ' search group(s) have no item and cannot be displayed.');
                try { loadMaster(rows, itemId); }
                catch (e) {
                    log.error({ title: 'Master CSV unavailable', details: e });
                    warnings.push('Master Excel Qty unavailable: ' + e.message);
                }
                fillItemNames(rows, warnings);
            } catch (e) {
                log.error({ title: 'Christmas quantity Suitelet failed', details: e });
                fatal = e.message || String(e);
            }
            const allRows = Array.from(rows.values());
            const selectedRows = allRows.filter(r => !chosenChannel || r.channel === chosenChannel);
            selectedRows.sort((a, b) => (a.name || a.id).localeCompare(b.name || b.id) || a.id.localeCompare(b.id) || a.channel.localeCompare(b.channel));
            if (!fatal && params.action === 'export' && !warnings.length) {
                context.response.writeFile({ file: file.create({ name: 'Christmas_Quantities.csv', fileType: file.Type.CSV, contents: exportCsv(selectedRows), encoding: file.Encoding.UTF8 }), isInline: false });
                return;
            }
            if (params.action === 'export' && warnings.length) warnings.push('CSV export is unavailable until the data errors above are resolved.');
            if (selectedRows.length > CONFIG.MAX_DISPLAY_ROWS) fatal = 'Too many rows to display. Narrow the filters or use Export CSV.';
            context.response.setHeader({ name: 'Content-Type', value: 'text/html; charset=UTF-8' });
            context.response.write(renderPage(selectedRows, allRows, itemId, chosenChannel, warnings, fatal));
        }
        return { onRequest };
    });
