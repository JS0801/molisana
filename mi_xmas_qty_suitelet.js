/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 *
 * Read-only Christmas quantities dashboard. Standalone CSV version.
 * No JSON configuration or third-party library is needed.
 */
define(['N/search', 'N/file', 'N/ui/serverWidget', 'N/log', 'N/runtime'],
    (search, file, serverWidget, log, runtime) => {
        'use strict';

        const CONFIG = {
            // File Cabinet folder containing master CSV uploads.
            MASTER_FOLDER_ID: 493530,
            // Newest upload = Created DESC, then Internal ID DESC.
            // Change to 'modified' only if overwriting existing files is preferred.
            MASTER_DATE_FIELD: 'created',
            MAX_FILE_BYTES: 9 * 1024 * 1024,
            MAX_MASTER_ROWS: 10000,
            MAX_DISPLAY_ROWS: 15000,
            SALES_ORDER_LOOKBACK_DAYS: 35
        };

        // LCL total is authoritative: do NOT add LCL/Fortinos detail again.
        // Metro total columns are deliberately omitted.
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
            return channelNames[key(text)] || text;
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
                id: itemId, name: clean(name), channel: canonical,
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
                ], 'AND', ['datecreated', 'within', 'daysago' + CONFIG.SALES_ORDER_LOOKBACK_DAYS, 'daysago0'], 'AND',
                ['taxline', 'is', 'F'], 'AND', ['cogs', 'is', 'F'], 'AND', ['closed', 'is', 'F']
            ];
            if (itemId) filters.push('AND', ['item', 'anyof', itemId]);
            const s = search.create({ type: 'salesorder', settings: [{ name: 'consolidationtype', value: 'ACCTTYPE' }], filters, columns });
            let withoutItem = 0;
            eachResult(s, r => {
                const id = r.getValue(columns[1]);
                if (!id || id === '- None -') { withoutItem++; return; }
                const row = getRow(rows, id, r.getText(columns[1]), r.getText(columns[5]));
                // Multiple customer categories collapse into the same item/channel.
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
                    ['folder', 'anyof', String(CONFIG.MASTER_FOLDER_ID)], 'AND',
                    ['name', 'endswith', '.csv'],
                    'AND', ['name', 'doesnotcontain', '~$']
                ], columns: cols
            }).run().getRange({ start: 0, end: 1 });
            if (!found.length) throw new Error('No .csv file found directly in master folder ' + CONFIG.MASTER_FOLDER_ID + '.');
            const f = file.load({ id: found[0].getValue(cols[1]) });
            if (f.size > CONFIG.MAX_FILE_BYTES) throw new Error('Newest master file exceeds the configured 9 MB limit: ' + f.name);
            return { file: f, date: found[0].getValue(cols[0]) };
        }

        // Comma-delimited CSV: quoted commas, escaped quotes and embedded
        // newlines are supported. Malformed quoting fails rather than shifts data.
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
            const required = ['Item ID'].concat(...MASTER_MAP.map(m => m.slice(1)));
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
                    const quantities = m.slice(1).map(h => {
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
            // Apply only after the entire file validates: never display a partial master.
            result.parsed.forEach(r => {
                if (!itemId || r.id === itemId) getRow(rows, r.id, '', r.channel).master = r.quantities;
            });
            return { name: newest.file.name, id: newest.file.id, date: newest.date, itemCount: result.itemCount };
        }

        function fillItemNames(rows, warnings) {
            const ids = Array.from(new Set(Array.from(rows.values()).filter(r => !r.name).map(r => r.id)));
            for (let start = 0; start < ids.length; start += 500) {
                const batch = ids.slice(start, start + 500);
                const names = new Map();
                eachResult(search.create({ type: 'item', filters: [['internalid', 'anyof', batch]], columns: [
                    search.createColumn({ name: 'internalid', sort: search.Sort.ASC }), 'itemid'
                ] }), r => { names.set(String(r.id), r.getValue({ name: 'itemid' })); });
                rows.forEach(r => { if (!r.name && names.has(r.id)) r.name = names.get(r.id); });
            }
            const unresolved = Array.from(new Set(Array.from(rows.values()).filter(r => !r.name).map(r => r.id)));
            if (unresolved.length) warnings.push(unresolved.length + ' master item ID(s) could not be resolved or are not visible to this role: ' + unresolved.slice(0, 20).join(', ') + '.');
        }

        function renderTable(rows, masterInfo, warnings) {
            const h = escapeHtml;
            const fmt = n => n === null ? '<span class="mi-missing" title="No saved value">—</span>' : h(n.toLocaleString('en-US', { maximumFractionDigits: 6 }));
            const group = (v, count, cls) => (v || Array(count).fill(null)).map(n =>
                '<td class="mi-num ' + cls + (n !== null && n < 0 ? ' mi-negative' : '') + '">' + fmt(n) + '</td>').join('');
            const body = rows.map(r => '<tr><td>' + h(r.id) + '</td><td class="mi-item">' + h(r.name || 'Unresolved item') + '</td><td>' + h(r.channel) + '</td>' +
                group(r.so, 3, 'mi-so') + group(r.reservation, 4, 'mi-or') + group(r.master, 3, 'mi-master') + '</tr>').join('');
            const notice = warnings.map(w => '<div class="mi-warning">' + h(w) + '</div>').join('');
            const source = masterInfo ? 'Master: ' + h(masterInfo.name) + ' (File ID ' + h(masterInfo.id) + ') · ' +
                h(CONFIG.MASTER_DATE_FIELD === 'created' ? 'Uploaded' : 'Modified') + ': ' + h(masterInfo.date) : 'Master Excel Qty is unavailable.';
            return `<style>
                .mi-dashboard{font:13px Arial,sans-serif;color:#172b4d;margin:16px 0;max-width:100%}
                .mi-meta{padding:12px 0;line-height:1.7}.mi-warning{background:#fff4d5;border-left:4px solid #d59b15;padding:12px;margin:8px 0;white-space:normal}
                .mi-scroll{overflow:auto;max-height:70vh;border:1px solid #d6dce5;border-radius:6px}
                .mi-dashboard table{border-collapse:separate;border-spacing:0;min-width:1300px;width:100%;font-size:12px}
                .mi-dashboard th,.mi-dashboard td{padding:10px 12px;border-right:1px solid #e2e7ed;border-bottom:1px solid #e2e7ed;white-space:nowrap}
                .mi-dashboard thead th{position:sticky;z-index:2;top:0;background:#263a55;color:white;text-align:center;height:24px}
                .mi-dashboard thead tr:nth-child(2) th{top:45px;background:#e8edf4;color:#25364c}
                .mi-dashboard thead .mi-head-so{background:#225e87}.mi-dashboard thead .mi-head-or{background:#4e568c}.mi-dashboard thead .mi-head-master{background:#247365}
                .mi-dashboard tbody tr:nth-child(even){background:#f7f9fc}.mi-dashboard tbody tr:hover{background:#edf4ff}
                .mi-dashboard .mi-num{text-align:right;font-variant-numeric:tabular-nums}.mi-dashboard .mi-item{min-width:180px}
                .mi-dashboard .mi-negative{color:#b52929}.mi-missing{color:#8c96a4}.mi-note{margin-top:10px;color:#526273;line-height:1.6}
                </style><div class="mi-dashboard">${notice}<div class="mi-meta"><strong>${rows.length.toLocaleString('en-US')} item/channel rows</strong><br>${source}</div>
                <div class="mi-scroll"><table aria-label="Christmas item quantities"><thead><tr>
                <th colspan="3">Item Details</th><th colspan="3" class="mi-head-so">Sales Order Qty</th><th colspan="4" class="mi-head-or">Order Reservation</th><th colspan="3" class="mi-head-master">Master Excel Qty</th>
                </tr><tr><th>Item ID</th><th>Item Name</th><th>Sales Channel</th><th>Ordered</th><th>Shipped</th><th>Remaining</th><th>Reserved</th><th>Committed</th><th>Used</th><th>Remaining</th><th>Booked</th><th>Filled</th><th>Balance</th></tr></thead>
                <tbody>${body || '<tr><td colspan="13">No matching item/channel rows.</td></tr>'}</tbody></table></div>
                <div class="mi-note">Sales orders: created within the last ${CONFIG.SALES_ORDER_LOOKBACK_DAYS} days, using the supplied Christmas criteria. Reservations: supplied open reservation criteria, without a date restriction.<br>
                LCL includes Fortinos. Blank sales channels map to Xmas 2026. Master values come from the latest CSV upload. — means no matching source row or a blank master value; zero means a recorded zero.</div></div>`;
        }

        function onRequest(context) {
            const params = context.request.parameters || {};
            let itemId = clean(params.custpage_item);
            const chosenChannel = clean(params.custpage_channel);
            const form = serverWidget.createForm({ title: 'MI | Christmas Quantity Comparison' });
            const itemField = form.addField({ id: 'custpage_item', type: serverWidget.FieldType.SELECT, label: 'Item', source: 'item' });
            const channelField = form.addField({ id: 'custpage_channel', type: serverWidget.FieldType.SELECT, label: 'Sales Channel' });
            form.addSubmitButton({ label: 'Apply Filters / Refresh' });
            const output = form.addField({ id: 'custpage_dashboard', type: serverWidget.FieldType.INLINEHTML, label: 'Quantities' });
            const rows = new Map(), warnings = [];
            let masterInfo = null, fatal = null;
            try {
                if (itemId && !/^\d+$/.test(itemId)) throw new Error('Select a valid item.');
                if (itemId) itemField.defaultValue = itemId;
                const omitted = loadSalesOrders(rows, itemId) + loadReservations(rows, itemId);
                if (omitted) warnings.push(omitted + ' search group(s) have no item and were excluded from the item comparison.');
                try { masterInfo = loadMaster(rows, itemId); }
                catch (e) {
                    log.error({ title: 'Master Excel unavailable', details: e });
                    warnings.push('Master Excel Qty unavailable: ' + e.message);
                }
                fillItemNames(rows, warnings);
            } catch (e) {
                log.error({ title: 'Christmas quantity Suitelet failed', details: e });
                fatal = e.message || String(e);
            }
            const channels = new Set(MASTER_MAP.map(m => m[0]));
            rows.forEach(r => channels.add(r.channel));
            if (chosenChannel) channels.add(channelName(chosenChannel));
            channelField.addSelectOption({ value: '', text: 'All Sales Channels' });
            Array.from(channels).sort((a, b) => a.localeCompare(b)).forEach(c => channelField.addSelectOption({ value: c, text: c }));
            if (chosenChannel) channelField.defaultValue = channelName(chosenChannel);
            let selectedRows = Array.from(rows.values()).filter(r => !chosenChannel || key(r.channel) === key(channelName(chosenChannel)));
            selectedRows.sort((a, b) => (a.name || a.id).localeCompare(b.name || b.id) || a.id.localeCompare(b.id) || a.channel.localeCompare(b.channel));
            if (selectedRows.length > CONFIG.MAX_DISPLAY_ROWS) fatal = 'More than ' + CONFIG.MAX_DISPLAY_ROWS + ' rows match. Select an item or sales channel and refresh. No rows have been silently truncated.';
            output.defaultValue = fatal ? '<div style="padding:18px;color:#9a2626;background:#fff2f2"><strong>Unable to display quantities.</strong><br>' + escapeHtml(fatal) + '</div>' : renderTable(selectedRows, masterInfo, warnings);
            context.response.writePage(form);
        }
        return { onRequest };
    });
