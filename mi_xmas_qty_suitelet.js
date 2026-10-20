/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * Read-only Christmas commitments and inventory dashboard.
 * Master bookings must use the same base units as NetSuite search quantities.
 * Configure inclusive season and purchase dates below before deployment.
 */
define(['N/search', 'N/file', 'N/log', 'N/runtime', 'N/format'],
    (search, file, log, runtime, format) => {
        'use strict';

        const CONFIG = {
            MASTER_FOLDER_ID: 493530,
            APPLICATION_ORIGIN: 'https://4975346.app.netsuite.com',
            MASTER_DATE_FIELD: 'created',
            MAX_FILE_BYTES: 9 * 1024 * 1024,
            MAX_MASTER_ROWS: 10000,
            MAX_DISPLAY_ROWS: 15000,
            // Inclusive dates. Update these for each Christmas season.
            SEASON_START: '2026-09-01',
            SEASON_END: '2026-12-31',
            PURCHASE_START: '2026-01-01'
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
                throw new Error('The full snapshot exceeds the remaining Suitelet execution budget. This section cannot be loaded completely; no partial totals are displayed.');
            }
        }
        function eachResult(s, callback) {
            const pages = s.runPaged({ pageSize: 1000 });
            if (pages.count > 1000000) throw new Error('Search exceeds the paging limit. A cached snapshot is required for this dataset.');
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

        function christmasOrderScope() {
            return [[
                    ['custbody_order_type_cp', 'anyof', '3'], 'OR', ['memomain', 'contains', 'Xmas'], 'OR',
                    ['formulatext: {otherrefnum}', 'contains', 'CHRISTMAS'], 'OR',
                    ['formulatext: {otherrefnum}', 'contains', 'Xmas'], 'OR',
                    ['memomain', 'contains', 'CHRISTMAS'], 'OR',
                    ["formulanumeric: CASE WHEN {item} LIKE '%XMAS%' OR {item} LIKE '%BACI%' THEN 1 ELSE 0 END", 'equalto', '1'], 'OR',
                    ["formulanumeric: CASE WHEN UPPER({item.displayname}) LIKE '%CHIOSTRO%' OR UPPER({item.displayname}) LIKE '%STREGA%' THEN 1 ELSE 0 END", 'equalto', '1']
                ], 'AND', [
                    ['actualshipdate', 'within', accountDate(CONFIG.SEASON_START), accountDate(CONFIG.SEASON_END)], 'OR',
                    ['trandate', 'within', accountDate(CONFIG.SEASON_START), accountDate(CONFIG.SEASON_END)], 'OR',
                    [['closed', 'is', 'F'], 'AND', ['formulanumeric: ABS({quantity}) - ABS(NVL({quantityshiprecv},0))', 'greaterthan', '0']]
                ]];
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
                christmasOrderScope(), 'AND', ['taxline', 'is', 'F'], 'AND', ['cogs', 'is', 'F']
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

        function loadReservations(rows, warnings) {
            const details=readSection('Order reservations',()=>{
                const out=[];
                // All-item version of the supplied item-6953 reservation search.
                // SUM(committed - allocated) equals SUM(committed) - SUM(allocated).
                // Read detail lines once, then group by item/channel in JavaScript.
                transactionRows('orderreservation', [
                    ['type','anyof','OrdResv'],'AND',['closed','is','F'],'AND',['mainline','is','F']
                ], ['item','tranid','saleschannel','location','quantity','quantitycommitted','quantitydemandallocated','quantityshiprecv'],r=>{
                    const itemId=String(r.getValue('item') || '');
                    if(!/^\d+$/.test(itemId)) return;
                    out.push({itemId,itemName:r.getText('item'),recordId:r.id,number:r.getValue('tranid'),channelId:String(r.getValue('saleschannel') || ''),channel:r.getText('saleschannel'),location:r.getText('location'),
                        reserved:numberValue(r.getValue('quantity'),'reservation quantity',0),
                        committed:numberValue(r.getValue('quantitycommitted'),'reservation committed',0)-numberValue(r.getValue('quantitydemandallocated'),'reservation allocated demand',0),
                        used:numberValue(r.getValue('quantitydemandallocated'),'reservation allocated demand',0),
                        fulfilled:numberValue(r.getValue('quantityshiprecv'),'reservation fulfilled/received',0)});
                },[{name:'consolidationtype',value:'ACCTTYPE'}]);
                return out;
            },warnings);
            if(details) details.forEach(r=>{
                const row=getRow(rows,r.itemId,r.itemName,r.channel);
                addQuantities(row,'reservation',[r.reserved,r.committed,r.used,r.fulfilled]);
            });
            return details;
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
            // NetSuite text filters do not support "endswith". Check the exact
            // extension in JavaScript, preserving newest-first order across pages.
            const pages = search.create({
                type: 'file',
                filters: [['folder', 'anyof', String(CONFIG.MASTER_FOLDER_ID)]],
                columns: cols
            }).runPaged({ pageSize: 1000 });
            let found = null;
            for (const page of pages.pageRanges) {
                assertUsage();
                found = pages.fetch({ index: page.index }).data.find(r => /\.csv$/i.test(String(r.getValue(cols[2]) || '')));
                if (found) break;
            }
            if (!found) throw new Error('No .csv file found directly in master folder ' + CONFIG.MASTER_FOLDER_ID + '.');
            const f = file.load({ id: found.getValue(cols[1]) });
            if (f.size > CONFIG.MAX_FILE_BYTES) throw new Error('Newest master file exceeds the configured 9 MB limit: ' + f.name);
            return { file: f, date: found.getValue(cols[0]) };
        }

        function parseCsv(contents, maxRows = CONFIG.MAX_MASTER_ROWS) {
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
                if (rows.length > maxRows + 1) throw new Error('CSV exceeds ' + maxRows + ' data rows.');
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

        // Missing or blank master bookings remain unknown, rather than becoming zero.
        // A missing sales group means no shipments or open demand matched the report scope.
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
                'Sales Order Qty - Ordered / demand', 'Sales Order Qty - Christmas SO Fulfilled', 'Sales Order Qty - Remaining',
                'Remaining Commitments', 'Commitment Status',
                'Order Reservation - Reserved', 'Order Reservation - Committed minus Allocated Demand', 'Order Reservation - Allocated Demand', 'Order Reservation - Fulfilled/Received'];
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
                '</tr><tr><th>Item ID</th><th>Item Description</th><th>Item Name</th><th>Sales Channel</th><th>Booked / Original Commitments</th><th>Ordered / demand</th><th>Christmas SO fulfilled</th><th>Remaining</th><th>Remaining Commitments</th><th>Status</th><th>Reserved</th><th>Committed − Allocated Demand</th><th>Allocated Demand</th><th>Fulfilled/Received</th></tr></thead><tbody>' +
                (rows.map(r => {
                    const commitment = commitmentSummary(r);
                    return '<tr class="' + commitment.css + '"><td>' + escapeHtml(r.id) + '</td><td class="description">' + escapeHtml(r.description || '—') + '</td><td>' + escapeHtml(r.name || 'Unresolved item') + '</td><td>' + escapeHtml(r.channel) + '</td>' +
                        group([commitment.booked], 1) + group(r.so || [0, 0, 0], 3) + group([commitment.remaining], 1) + '<td class="status">' + escapeHtml(commitment.status) + '</td>' +
                        group(r.reservation, 4) + '</tr>';
                }).join('') || '<tr><td colspan="14">No matching items.</td></tr>') + '</tbody></table></div>';
        }

        function accountDate(iso) {
            if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) throw new Error('Configure dates as YYYY-MM-DD.');
            const [y, m, d] = iso.split('-').map(Number);
            const date = new Date(y, m - 1, d);
            if(date.getFullYear()!==y || date.getMonth()!==m-1 || date.getDate()!==d) throw new Error('Invalid configured date: '+iso);
            return format.format({ value: date, type: format.Type.DATE });
        }
        const quantity = r => Math.abs(numberValue(r, 'quantity', 0));
        function lineFilters(ids) {
            return [['mainline', 'is', 'F'], 'AND', ['taxline', 'is', 'F'], 'AND', ['cogs', 'is', 'F'], 'AND', ['shipping', 'is', 'F'], 'AND', ['item', 'anyof', ids]];
        }
        function transactionRows(type, filters, columns, consume, settings) {
            const order = [search.createColumn({name:'internalid', sort:search.Sort.ASC}), search.createColumn({name:'lineuniquekey', sort:search.Sort.ASC})];
            const options={type,filters,columns:order.concat(columns)};
            if(settings)options.settings=settings;
            eachResult(search.create(options), consume);
        }
        // Load into temporary arrays. A failed later page must not publish partial totals.
        function readSection(label, fn, warnings) {
            try { return fn(); }
            catch (e) {
                log.error({title:label + ' unavailable', details:e});
                warnings.push(label + ' unavailable: ' + (e.message || String(e)));
                return null;
            }
        }
        function newInventory(id) {
            return {id, stock:null, purchases:null, openOrders:null, shipments:null, transfers:null, reservations:null};
        }
        function loadInventory(rows, itemId, warnings, reservationData) {
            const ids = Array.from(new Set(Array.from(rows.values()).map(r=>r.id)));
            const inventory = new Map(ids.map(id=>[id,newInventory(id)]));
            // Clear the discovery quantities; fulfilled quantity comes from Christmas SO lines, including billed orders.
            rows.forEach(r=>{r.so=null;});
            if (!ids.length) return inventory;
            {
                const batch=ids; // One search per dataset, not one search per item/batch.
                const stock=readSection('Stock by location',()=>{
                    const out=[];
                    eachResult(search.create({type:'item', filters:[['internalid','anyof',batch]], columns:[
                        search.createColumn({name:'internalid',sort:search.Sort.ASC}), search.createColumn({name:'inventorylocation',sort:search.Sort.ASC}),
                        'itemid','displayname','salesdescription','locationquantityonhand','locationquantityavailable','locationquantitycommitted','locationquantityonorder','locationquantitybackordered','locationquantityintransit'
                    ]}),r=>out.push({id:String(r.id),name:clean(r.getValue('itemid')),description:clean(r.getValue('salesdescription')) || clean(r.getValue('displayname')),location:r.getText('inventorylocation') || 'No inventory location',
                        onHand:numberValue(r.getValue('locationquantityonhand'),'on hand',0), available:numberValue(r.getValue('locationquantityavailable'),'available',0),
                        committed:quantity(r.getValue('locationquantitycommitted')), onOrder:quantity(r.getValue('locationquantityonorder')),
                        backordered:quantity(r.getValue('locationquantitybackordered')), inTransit:quantity(r.getValue('locationquantityintransit'))}));
                    return out;
                },warnings);
                const purchases=readSection('Purchase orders',()=>{
                    const out=[];
                    transactionRows('purchaseorder',lineFilters(batch).concat(['AND',['status','noneof','PurchOrd:A','PurchOrd:C'],'AND',[['trandate','within',accountDate(CONFIG.PURCHASE_START),accountDate(CONFIG.SEASON_END)],'OR',[['closed','is','F'],'AND',['formulanumeric: ABS({quantity}) - ABS(NVL({quantityshiprecv},0))','greaterthan','0']]]]),
                        ['item','tranid','trandate','entity','location','quantity','quantityshiprecv','closed','statusref','expectedreceiptdate'],r=>{
                            const ordered=quantity(r.getValue('quantity')),received=quantity(r.getValue('quantityshiprecv'));
                            const closed=r.getValue('closed')===true || r.getValue('closed')==='T' || r.getValue('statusref')==='closed';
                            out.push({id:String(r.getValue('item')),recordId:r.id,number:r.getValue('tranid'),date:r.getValue('trandate'),vendor:r.getText('entity'),location:r.getText('location'),ordered,received,
                                remaining:closed?0:Math.max(0,ordered-received),expected:r.getValue('expectedreceiptdate')});
                        }); return out;
                },warnings);
                const openOrders=readSection('Open sales orders',()=>{
                    const out=[];
                    transactionRows('salesorder',lineFilters(batch).concat(['AND',['closed','is','F'],'AND',['status','noneof','SalesOrd:C','SalesOrd:H'],'AND',
                        ['formulanumeric: ABS({quantity}) - ABS(NVL({quantityshiprecv},0))','greaterthan','0']]),
                        ['item','tranid','entity','location','quantity','quantityshiprecv','quantitycommitted',search.createColumn({name:'custentity_default_shipping_channel',join:'customerMain'})],r=>{
                            const ordered=quantity(r.getValue('quantity')),fulfilled=quantity(r.getValue('quantityshiprecv')),committed=quantity(r.getValue('quantitycommitted'));
                            out.push({id:String(r.getValue('item')),recordId:r.id,number:r.getValue('tranid'),customer:r.getText('entity'),location:r.getText('location'),ordered,fulfilled,committed,
                                remaining:Math.max(0,ordered-fulfilled),uncommitted:Math.max(0,ordered-fulfilled-committed),channel:channelName(r.getText({name:'custentity_default_shipping_channel',join:'customerMain'}))});
                        });return out;
                },warnings);
                const shipments=readSection('Christmas sales order fulfillments',()=>{
                    const out=[];
                    // All statuses are eligible: billed, pending billing, partial and closed
                    // lines can have fulfilled quantities. Never add open-order fulfilled
                    // quantities again; the open search contributes remaining demand only.
                    transactionRows('salesorder',lineFilters(batch).concat(['AND',christmasOrderScope(),'AND',
                        ['formulanumeric: ABS(NVL({quantityshiprecv},0))','greaterthan','0']]),
                        ['item','quantityshiprecv',search.createColumn({name:'custentity_default_shipping_channel',join:'customerMain'})],r=>out.push({id:String(r.getValue('item')),shipped:quantity(r.getValue('quantityshiprecv')),
                            channel:channelName(r.getText({name:'custentity_default_shipping_channel',join:'customerMain'}))})); return out;
                },warnings);
                for (const [field,data] of Object.entries({stock,purchases,openOrders,shipments})) distribute(inventory,field,data,'id');
                inventory.forEach(inv=>{
                    if(inv.stock && !inv.stock.length) {inv.stock=null;warnings.push('Stock unavailable for item '+inv.id+'.');}
                });
                if (openOrders) openOrders.forEach(o=>{
                    const row=getRow(rows,o.id,'',o.channel);
                    if (!row.so) row.so=[0,0,0];
                    row.so[0]+=o.remaining; row.so[2]+=o.remaining;
                });
                if (shipments) shipments.forEach(o=>{
                    const row=getRow(rows,o.id,'',o.channel);
                    if (!row.so) row.so=[0,0,0];
                    row.so[0]+=o.shipped; row.so[1]+=o.shipped;
                });
                rows.forEach(r=>{if(inventory.has(r.id)) {
                    const item=inventory.get(r.id).stock?.[0];
                    if(item) {r.name=item.name || r.name;r.description=item.description;}
                    if(!r.so) r.so=[0,0,0];
                    if(shipments===null) {r.so[0]=null;r.so[1]=null;}
                    if(openOrders===null) {r.so[0]=null;r.so[2]=null;}
                }});
            }
            loadAllDetails(inventory,warnings,reservationData);
            return inventory;
        }
        // Index each master dataset in one pass, then assign lists in memory.
        function distribute(inventory,field,data,idField) {
            inventory.forEach(inv=>{inv[field]=data===null?null:[];});
            if(data) data.forEach(row=>{const inv=inventory.get(String(row[idField]));if(inv)inv[field].push(row);});
        }
        function loadAllDetails(inventory,warnings,reservationData) {
            const ids=Array.from(inventory.keys());
            if(!ids.length)return;
            distribute(inventory,'reservations',reservationData===undefined?null:reservationData,'itemId');
            const allocations=readSection('Supply allocation details',()=>{
                const out=[];
                transactionRows('transaction',lineFilters(ids).concat(['AND',['type','anyof','SalesOrd','TrnfrOrd'],'AND',['closed','is','F'],'AND',
                    ['formulanumeric: ABS(NVL({quantityallocated},0))','greaterthan','0']]),['item','tranid','type','entity','location','quantityallocated'],r=>out.push({itemId:String(r.getValue('item')),recordId:r.id,number:r.getValue('tranid'),type:r.getText('type'),customer:r.getText('entity'),location:r.getText('location'),allocated:quantity(r.getValue('quantityallocated'))}));return out;
            },warnings);
            distribute(inventory,'allocations',allocations,'itemId');
            const transfers=readSection('Incoming transfer orders',()=>{
                const out=[],seen=new Map();
                // Outgoing (negative-quantity) side only: a transfer also has an incoming
                // side, which must not double the quantity. Dedicated transfer search
                // columns expose both shipped and received amounts without record loads.
                transactionRows('transferorder',lineFilters(ids).concat(['AND',['closed','is','F'],'AND',['status','noneof','TrnfrOrd:A','TrnfrOrd:C','TrnfrOrd:H'],'AND',['quantity','lessthan','0']]),
                    ['item','tranid','location','transferlocation','quantity','transferorderitemline',
                     'transferorderquantitycommitted','transferorderquantityshipped','transferorderquantityreceived','expectedreceiptdate'],r=>{
                        const itemId=String(r.getValue('item'));
                        if(!inventory.has(itemId))return;
                        const line=String(r.getValue('transferorderitemline') || r.getValue('lineuniquekey') || '');
                        if(!line)throw new Error('Transfer '+r.id+' has no usable source line ID; refusing to merge different item lines.');
                        const ordered=quantity(r.getValue('quantity'));
                        const shipped=quantity(r.getValue('transferorderquantityshipped'));
                        const received=quantity(r.getValue('transferorderquantityreceived'));
                        const committed=quantity(r.getValue('transferorderquantitycommitted'));
                        const key=String(r.id)+'|'+line+'|'+itemId;
                        const values=[ordered,shipped,received,committed].join('|');
                        if(seen.has(key)) {
                            if(seen.get(key)!==values)throw new Error('Conflicting transfer quantities for '+r.id+', line '+line+'.');
                            return;
                        }
                        seen.set(key,values);
                        if(ordered<=received)return;
                        out.push({itemId,recordId:r.id,number:r.getValue('tranid'),from:r.getText('location'),to:r.getText('transferlocation'),units:'Base units',
                            ordered,shipped,received,committed,expected:r.getValue('expectedreceiptdate') || '',remaining:Math.max(0,ordered-received),inTransit:Math.max(0,shipped-received)});
                    });
                return out;
            },warnings);
            distribute(inventory,'transfers',transfers,'itemId');
        }
        function sum(data,field) {return data===null?null:data.reduce((s,r)=>s+r[field],0);}
        function summarizeInventory(inv,rows) {
            const bookings=rows.filter(r=>r.id===inv.id);
            const known=bookings.length>0 && bookings.every(r=>r.master && r.master[0]!==null && r.master[0]!==undefined);
            const booked=known?bookings.reduce((s,r)=>s+r.master[0],0):null;
            const shipped=sum(inv.shipments,'shipped'),onHand=sum(inv.stock,'onHand'),transit=sum(inv.stock,'inTransit'),incoming=sum(inv.purchases,'remaining');
            // Overshipments for customer A do not cancel a commitment still owed to customer B.
            const outstanding=known && shipped!==null?bookings.reduce((s,r)=>s+Math.max(0,commitmentSummary(r).remaining),0):null;
            const excess=outstanding===null || onHand===null || transit===null || incoming===null?null:onHand+transit+incoming-outstanding;
            const purchaseStart=accountDate(CONFIG.PURCHASE_START),purchaseEnd=accountDate(CONFIG.SEASON_END);
            const from=format.parse({value:purchaseStart,type:format.Type.DATE}).getTime(),to=format.parse({value:purchaseEnd,type:format.Type.DATE}).getTime();
            const seasonal=inv.purchases===null?null:inv.purchases.filter(p=>{const date=format.parse({value:p.date,type:format.Type.DATE}).getTime();return date>=from && date<=to;});
            return {id:inv.id,name:bookings[0]?.name || inv.id,description:bookings[0]?.description || '',booked,shipped,outstanding,onHand,
                available:sum(inv.stock,'available'),committed:sum(inv.stock,'committed'),onOrder:sum(inv.stock,'onOrder'),backordered:sum(inv.stock,'backordered'),transit,incoming,
                purchased:sum(seasonal,'ordered'),received:sum(seasonal,'received'),openDemand:sum(inv.openOrders,'remaining'),excess};
        }
        function detailTable(title,heads,data,values) {
            if(data===null || data===undefined) return '<h3>'+escapeHtml(title)+'</h3><p class="error">Data unavailable. See warnings above.</p>';
            const tx=r=>'<a target="_blank" rel="noopener" href="'+escapeHtml(CONFIG.APPLICATION_ORIGIN)+'/app/accounting/transactions/transaction.nl?id='+encodeURIComponent(r.recordId)+'">'+escapeHtml(r.number || r.recordId)+'</a>';
            return '<details open><summary>'+escapeHtml(title)+'</summary><div class="table-wrap detail"><table><thead><tr>'+heads.map(h=>'<th>'+escapeHtml(h)+'</th>').join('')+'</tr></thead><tbody>'+
                (data.map(r=>'<tr>'+values(r).map((v,i)=>'<td>'+(i===0 && r.recordId?tx(r):escapeHtml(v==null?'—':v))+'</td>').join('')+'</tr>').join('') || '<tr><td colspan="'+heads.length+'">No matching records.</td></tr>')+'</tbody></table></div></details>';
        }
        function renderItemDetails(inventory,itemId,summary) {
            let html = '';
            if(itemId && inventory.has(itemId)) {
                const i=inventory.get(itemId);
                html+='<div class="inline-details"><h3>Item Inventory — All Sales Channels</h3>';
                if(summary) html+='<dl class="inventory-grid">'+INVENTORY_FIELDS.map((field,index)=>'<div><dt>'+escapeHtml(INVENTORY_HEADINGS[index])+'</dt><dd class="'+(field==='excess' && summary[field]<0?'negative':'')+'">'+(summary[field]==null?'—':escapeHtml(summary[field].toLocaleString('en-US',{maximumFractionDigits:6})))+'</dd></div>').join('')+'</dl>';
                html+=detailTable('Stock by location',['Location','On hand','Available','Committed','On order','Backordered','Internal in transit'],i.stock,r=>[r.location,r.onHand,r.available,r.committed,r.onOrder,r.backordered,r.inTransit]);
                html+=detailTable('Open sales orders / stock holders',['Order','Customer','Location','Ordered','Fulfilled','Remaining','Committed','Uncommitted'],i.openOrders,r=>[r.number,r.customer,r.location,r.ordered,r.fulfilled,r.remaining,r.committed,r.uncommitted]);
                html+=detailTable('Supply allocations (do not add to committed quantities)',['Order','Type','Customer','Location','Allocated'],i.allocations,r=>[r.number,r.type,r.customer,r.location,r.allocated]);
                html+=detailTable('Open order reservations',['Reservation','Channel','Location','Reserved','Committed − Allocated Demand','Allocated Demand','Fulfilled/Received'],i.reservations,r=>[r.number,r.channel,r.location,r.reserved,r.committed,r.used,r.fulfilled]);
                html+=detailTable('Incoming purchase orders',['PO','Vendor','Location','Ordered','Received','Remaining','Expected receipt'],i.purchases===null?null:i.purchases.filter(p=>p.remaining>0),r=>[r.number,r.vendor,r.location,r.ordered,r.received,r.remaining,r.expected]);
                html+=detailTable('Incoming internal transfers (base units)',['Transfer','From','To','Units','Ordered','Committed','Dispatched','Received','Remaining to receive','In transit','Expected receipt'],i.transfers,r=>[r.number,r.from,r.to,r.units,r.ordered,r.committed,r.shipped,r.received,r.remaining,r.inTransit,r.expected]);
                html+='</div>';
            }
            return html;
        }
        const INVENTORY_FIELDS = ['purchased','received','onHand','available','committed','onOrder','backordered','transit','incoming','shipped','booked','outstanding','openDemand','excess'];
        const INVENTORY_HEADINGS = ['Purchased','Received against POs','On hand','Available','Committed','On order (reference)','Backordered','Internal transfer in transit','Open PO remaining','Christmas SO fulfilled — all channels','Booked — all channels','Outstanding commitments — all channels','Open SO demand — all channels','Projected excess / (shortage)'];
        function renderCombined(rows, allRows, inventory, chosenChannel, summaries) {
            const groups = new Map();
            rows.forEach(r => {if(!groups.has(r.id)) groups.set(r.id,[]);groups.get(r.id).push(r);});
            const fmt = n => n == null ? '—' : escapeHtml(n.toLocaleString('en-US',{maximumFractionDigits:6}));
            const numbers = values => values.map(n=>'<td class="num'+(n<0?' negative':'')+'">'+fmt(n)+'</td>').join('');
            const heads=['Details','Item','Sales Channel','Booked / Original Commitments','Ordered / demand','Christmas SO fulfilled','Open SO remaining','Remaining Commitments','Status','Reserved','Committed − Allocated Demand','Allocated Demand','Fulfilled/Received'];
            let body='';
            const csvIndexes=new Map(rows.map((r,i)=>[r,i]));
            groups.forEach((channels,id)=>{
                const inv=inventory.get(id);
                const summary=inv ? (summaries ? summaries.get(id) : summarizeInventory(inv,allRows)) : null;
                body+='<tbody class="item-group" data-item-id="'+escapeHtml(id)+'" data-search="'+escapeHtml([id,channels[0].name,channels[0].description].join(' '))+'" data-expanded="false">';
                channels.forEach((r,index)=>{
                    const commitment=commitmentSummary(r);
                    body+='<tr class="channel-row '+commitment.css+'" data-channel="'+escapeHtml(r.channel)+'" data-status="'+escapeHtml(commitment.status)+'" data-csv-index="'+csvIndexes.get(r)+'">';
                    if(index===0) {
                        body+='<td class="item-total shared-leading" rowspan="'+channels.length+'"><button type="button" class="toggle-detail" aria-expanded="false" aria-controls="detail-'+escapeHtml(id)+'">View details</button></td>';
                        body+='<td class="description item-total shared-leading" rowspan="'+channels.length+'"><strong>'+escapeHtml(r.name || id)+'</strong><div>'+escapeHtml(r.description || '—')+'</div><small>Item ID: '+escapeHtml(id)+'</small></td>';
                    }
                    body+='<td>'+escapeHtml(r.channel)+'</td>'+numbers([commitment.booked])+numbers(r.so || [0,0,0])+numbers([commitment.remaining])+'<td class="status">'+escapeHtml(commitment.status)+'</td>'+numbers(r.reservation || [null,null,null,null]);
                    body+='</tr>';
                });
                body+='<tr id="detail-'+escapeHtml(id)+'" class="item-detail-row" hidden><td colspan="'+heads.length+'">'+'<template class="detail-template">'+renderItemDetails(inventory,id,summary)+'</template></td></tr></tbody>';
            });
            return '<div class="table-wrap"><table id="combined-table"><thead><tr><th>Details</th><th colspan="2">Item / Sales Channel</th><th class="master">Original Booking</th><th colspan="3" class="so">Sales Order Qty</th><th colspan="2" class="master">Commitment Progress</th><th colspan="4" class="reservation">Order Reservation</th></tr><tr>'+heads.map(h=>'<th>'+escapeHtml(h)+'</th>').join('')+'</tr></thead>'+body+'<tbody id="empty-result" hidden><tr><td colspan="'+heads.length+'">No matching items.</td></tr></tbody></table></div>';
        }

        function exportCombined(rows,allRows,inventory) {
            const table=parseCsv(exportCsv(rows), CONFIG.MAX_DISPLAY_ROWS);
            table[0]=table[0].concat(INVENTORY_HEADINGS);
            const seen=new Set();
            rows.forEach((r,index)=>{
                const summary=!seen.has(r.id) && inventory.has(r.id)?summarizeInventory(inventory.get(r.id),allRows):null;
                table[index+1]=table[index+1].concat(INVENTORY_FIELDS.map(f=>summary ? summary[f] : null));
                seen.add(r.id);
            });
            // Existing text values have already been spreadsheet-escaped by exportCsv.
            const cell=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
            return '\uFEFF'+table.map(r=>r.map(cell).join(',')).join('\r\n');
        }

        function exportInventory(inventory,allRows,visibleRows) {
            const visible=new Set(visibleRows.map(r=>r.id));
            const fields=['id','description','name','purchased','received','onHand','available','committed','onOrder','transit','incoming','shipped','booked','outstanding','openDemand','excess'];
            const cell=v=>{let s=v==null?'':String(v);if(typeof v==='string' && /^\s*[=+@-]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';};
            return '\uFEFF'+[fields].concat(Array.from(inventory.values()).filter(i=>visible.has(i.id)).map(i=>{const summary=summarizeInventory(i,allRows);return fields.map(f=>summary[f]);})).map(r=>r.map(cell).join(',')).join('\r\n');
        }

        // Runs entirely in the browser. No fetch, navigation or form submission.
        function browserController() {
            const data=JSON.parse(document.getElementById('report-data').textContent);
            const form=document.getElementById('report-filters');
            const item=document.getElementById('item-filter');
            const status=document.getElementById('status-filter');
            const channel=document.getElementById('channel-filter');
            const groups=Array.from(document.querySelectorAll('tbody.item-group'));
            const normalize=v=>String(v || '').trim().toLowerCase();
            const query=document.getElementById('filter-status');
            // Cache main-table nodes; never traverse all hidden details on each keystroke.
            const groupNodes=new Map(groups.map(group=>[group,{
                rows:Array.from(group.querySelectorAll('tr.channel-row')),
                detail:group.querySelector('tr.item-detail-row'),
                leading:Array.from(group.querySelectorAll('td.shared-leading')),
                trailing:Array.from(group.querySelectorAll('td.shared-trailing')),
                search:normalize(group.dataset.search)
            }]));
            const searchableFilters=[item,channel,status].map(input=>{
                const list=document.getElementById(input.id+'-options');
                const options=Array.from(list.querySelectorAll('option')).map(o=>({value:o.value,id:o.dataset.id || o.value}));
                function refresh() {
                    const term=normalize(input.value);
                    const matches=options.filter(o=>!term || normalize(o.value).includes(term));
                    matches.sort((a,b)=>Number(!normalize(a.value).startsWith(term))-Number(!normalize(b.value).startsWith(term)) || a.value.localeCompare(b.value));
                    list.replaceChildren(...matches.map(o=>{
                        const option=document.createElement('option');option.value=o.value;option.dataset.id=o.id;return option;
                    }));
                }
                input.addEventListener('input',()=>{refresh();applyFilters();});
                input.addEventListener('change',applyFilters);
                return {input,options,refresh};
            });
            function applyFilters() {
                const term=normalize(item.value),selected=normalize(channel.value),selectedStatus=normalize(status.value);
                const exactItem=searchableFilters[0].options.find(o=>normalize(o.value)===term || o.id===item.value.trim());
                const selectedItem=exactItem?exactItem.id:'';
                let visibleItems=0,visibleChannels=0;
                groups.forEach(group=>{
                    const matchesItem=selectedItem ? group.dataset.itemId===selectedItem : (!term || groupNodes.get(group).search.includes(term));
                    const cached=groupNodes.get(group),rows=cached.rows;
                    const visible=rows.filter(row=>matchesItem && (!selected || normalize(row.dataset.channel).includes(selected)) && (!selectedStatus || normalize(row.dataset.status).includes(selectedStatus)));
                    const visibleSet=new Set(visible);
                    rows.forEach(row=>{row.hidden=!visibleSet.has(row);});
                    group.hidden=visible.length===0;
                    const detail=cached.detail;
                    detail.hidden=!visible.length || group.dataset.expanded!=='true';
                    if(!visible.length) return;
                    visibleItems++;visibleChannels+=visible.length;
                    // Move shared cells to the first visible channel, preserving alignment
                    // when the originally first channel has been filtered out.
                    const first=visible[0];
                    const leading=cached.leading,trailing=cached.trailing;
                    leading.forEach(cell=>{cell.rowSpan=visible.length;});
                    for(let n=leading.length-1;n>=0;n--)first.insertBefore(leading[n],first.firstChild);
                    trailing.forEach(cell=>{cell.rowSpan=visible.length;first.appendChild(cell);});
                });
                document.getElementById('empty-result').hidden=visibleItems!==0;
                query.textContent=visibleItems+' item(s), '+visibleChannels+' sales-channel row(s). showing.';
            }
            form.addEventListener('submit',event=>{event.preventDefault();applyFilters();});
            document.getElementById('clear-filters').addEventListener('click',()=>{searchableFilters.forEach(f=>{f.input.value='';f.refresh();});applyFilters();});
            document.getElementById('combined-table').addEventListener('click',event=>{
                const button=event.target.closest('button.toggle-detail');
                if(!button)return;
                const group=button.closest('tbody.item-group');
                const expanded=group.dataset.expanded!=='true';
                // Details were delivered in the initial HTML but remain inert until opened.
                const template=group.querySelector('template.detail-template');
                if(expanded && template) {
                    template.parentNode.appendChild(template.content.cloneNode(true));
                    template.remove();
                }
                group.dataset.expanded=String(expanded);
                group.querySelector('tr.item-detail-row').hidden=!expanded;
                button.setAttribute('aria-expanded',String(expanded));
                button.textContent=expanded?'Hide details':'View details';
            });
            document.getElementById('export-csv').addEventListener('click',()=>{
                if(!data.canExport)return;
                const table=[data.headers];
                groups.filter(g=>!g.hidden).forEach(group=>{
                    const visible=Array.from(group.querySelectorAll('tr.channel-row')).filter(r=>!r.hidden);
                    visible.forEach((row,index)=>{
                        const cells=data.rows[Number(row.dataset.csvIndex)].slice();
                        table.push(cells.concat(index===0?data.inventory[group.dataset.itemId]:Array(data.inventoryColumns).fill(null)));
                    });
                });
                // Channel text has already been escaped for spreadsheet formulas on the server.
                const cell=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
                const csv='\uFEFF'+table.map(row=>row.map(cell).join(',')).join('\r\n');
                const objectUrl=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
                const link=document.createElement('a');link.href=objectUrl;link.download='Christmas_Quantities.csv';
                document.body.appendChild(link);link.click();link.remove();
                setTimeout(()=>URL.revokeObjectURL(objectUrl),1000);
            });
            applyFilters();
        }
        function renderBrowserScript(rows,inventory,warnings,summaries) {
            const csv=parseCsv(exportCsv(rows), CONFIG.MAX_DISPLAY_ROWS);
            const totals={};
            inventory.forEach((inv,id)=>{const summary=summaries ? summaries.get(id) : summarizeInventory(inv,rows);totals[id]=INVENTORY_FIELDS.map(f=>summary[f]);});
            const data={headers:csv[0].concat(INVENTORY_HEADINGS),rows:csv.slice(1),inventory:totals,inventoryColumns:INVENTORY_FIELDS.length,canExport:!warnings.length};
            const json=JSON.stringify(data).replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
            return '<script type="application/json" id="report-data">'+json+'</script><script>('+browserController.toString()+')();</script>';
        }

        function renderPage(rows, allRows, itemId, chosenChannel, warnings, fatal, inventory) {
            // Index once and reuse each summary for the table and CSV payload.
            const rowsByItem=new Map(),summaries=new Map();
            allRows.forEach(r=>{if(!rowsByItem.has(r.id))rowsByItem.set(r.id,[]);rowsByItem.get(r.id).push(r);});
            inventory.forEach((inv,id)=>summaries.set(id,summarizeInventory(inv,rowsByItem.get(id) || [])));
            const h = escapeHtml;
            const items = new Map();
            allRows.forEach(r => items.set(r.id, [r.name || r.id, r.description].filter(Boolean).join(' — ')));
            const options = Array.from(items).sort((a,b)=>a[1].localeCompare(b[1])).map(([id,name])=>
                '<option data-id="'+h(id)+'" value="'+h(name+' (ID: '+id+')')+'"></option>').join('');
            const channels = MASTER_MAP.map(m=>m[0]).sort((a,b)=>a.localeCompare(b)).map(c=>
                '<option value="'+h(c)+'"></option>').join('');
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
                #item-filter{max-width:430px}#report-explanation{background:#fff;border:1px solid #dce2eb;padding:16px;margin:16px 0;line-height:1.6}.inventory-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;white-space:normal;margin:16px 0}.inventory-grid>div{background:#fff;border:1px solid #dce2eb;border-radius:6px;padding:12px}.inventory-grid dt{font-size:12px;color:#4b5e73}.inventory-grid dd{margin:8px 0 0;font-size:20px;font-weight:700}.description strong,.description small{display:block}.description small{margin-top:6px;color:#56677b}
                .legend{margin:12px 0;line-height:1.6;color:#394c63}
                [hidden]{display:none!important}.item-detail-row>td{background:#f6f8fb;padding:18px}.inline-details{text-align:left;white-space:normal}.inline-details h3{margin:8px 0}button:disabled{opacity:.5;cursor:not-allowed}
                td.item-total{background:#f0f4f8;vertical-align:top}td.shortage{background:#fde5e5;color:#a5202d;font-weight:bold}
                details{margin:18px 0}summary{font-size:16px;font-weight:bold;cursor:pointer}.detail{max-height:420px}.detail table{min-width:900px}.table-wrap form{padding:0;border:0;background:transparent}h2{margin-top:28px}
                @media(max-width:650px){main{padding:14px}label,input,select{width:100%}h1{font-size:21px}}
                </style></head><body><main><h1>Christmas Quantity Comparison</h1>
                <form id="report-filters" onsubmit="return false">
                <label for="item-filter">Item<input id="item-filter" list="item-filter-options" placeholder="All items — type ID, code or description" autocomplete="off"><datalist id="item-filter-options">${options}</datalist></label>
                <label for="channel-filter">Sales Channel<input id="channel-filter" list="channel-filter-options" placeholder="All channels — type to search" autocomplete="off"><datalist id="channel-filter-options">${channels}</datalist></label>
                <label for="status-filter">Status<input id="status-filter" list="status-filter-options" placeholder="All statuses — type to search" autocomplete="off"><datalist id="status-filter-options">${['Completed','Partially fulfilled','Over commitment','Not shipped','No commitment','Booking unavailable'].map(s=>'<option value="'+h(s)+'"></option>').join('')}</datalist></label>
                <button type="submit">Apply Filters</button><button type="button" id="export-csv" class="secondary" ${warnings.length || fatal ? 'disabled title="Resolve data warnings before exporting"' : ''}>Export CSV</button>
                <button type="button" id="clear-filters" class="secondary">Clear Filters</button>
                <button type="button" class="secondary" aria-expanded="false" aria-controls="report-explanation" onclick="var p=document.getElementById('report-explanation');p.hidden=!p.hidden;this.setAttribute('aria-expanded',String(!p.hidden));">Report Details</button></form>
                <section id="report-explanation" hidden>
                <p class="note">Fulfilled quantities are cumulative on Christmas sales orders with an order date or actual ship date from ${h(CONFIG.SEASON_START)} through ${h(CONFIG.SEASON_END)}, plus older open Christmas orders. Billed, partially fulfilled, pending-billing and closed orders with fulfillment are included. These are SO fulfilled quantities, not shipments restricted to fulfillment dates. Purchases: ${h(CONFIG.PURCHASE_START)} through ${h(CONFIG.SEASON_END)}. Incoming supply and open demand use current open records, including older orders. “Ordered / demand” = Christmas SO fulfilled + current open SO quantity. Remaining Commitments = Original Booked − Christmas SO Fulfilled. Other channels are merged into Xmas 2026.</p>
                <p class="legend">Green: completed · Yellow: partially fulfilled · Red: shipped over commitment. Missing bookings show as unavailable; zero bookings with shipments show as over commitment. Order Reservation quantities are shown separately. Allocated Demand uses quantitydemandallocated. The adjusted committed column is quantitycommitted minus quantitydemandallocated. Fulfilled/Received uses quantityshiprecv. All open reservation lines are included, without a positive-remainder or sales-channel exclusion. Other channels are grouped into Xmas 2026.</p>
                <p>Inventory figures include all sales channels and visible locations, even when filters are applied. Master bookings must use item base units. Blank bookings remain unknown.</p>
                <p>Projected excess / shortage = On hand + internal transfer in transit + open PO remaining − outstanding commitments. On order is reference only. Supplier stock in transit is already included in open PO remaining. This is coverage against bookings, not free-to-promise stock.</p>
                <p>Filters and details use the loaded snapshot. Reload the page to refresh NetSuite data. Each filter is one searchable box. Type to filter immediately or choose a suggestion. Clear a box to show all values. Clear Filters restores every option. Export CSV includes filtered channel rows and item inventory totals once per item.</p>
                </section>
                ${warnings.map(w => '<div class="error" role="alert">' + h(w) + '</div>').join('')}
                ${fatal ? '<div class="error" role="alert">' + h(fatal) + '</div>' : '<p id="filter-status" role="status"></p>' + renderCombined(rows, allRows, inventory, chosenChannel, summaries) + renderBrowserScript(rows, inventory, warnings, summaries)}
                </main></body></html>`;
        }

        function onRequest(context) {
            // Always load the complete snapshot. URL/form filter parameters never narrow
            // server queries; filtering, details and CSV export are browser-only.
            const rows = new Map(), warnings = [];
            let inventory = new Map(), fatal = null;
            try {
                const omitted = loadSalesOrders(rows, '');
                if (omitted) warnings.push(omitted + ' search group(s) have no item and cannot be displayed.');
                try { loadMaster(rows, ''); }
                catch (e) {
                    log.error({ title: 'Master CSV unavailable', details: e });
                    warnings.push('Master Excel Qty unavailable: ' + e.message);
                }
                const reservationData = loadReservations(rows,warnings);
                if (rows.size > CONFIG.MAX_DISPLAY_ROWS) throw new Error('The full report exceeds CONFIG.MAX_DISPLAY_ROWS. Increase the configured limit after reviewing account capacity.');
                inventory = loadInventory(rows, '', warnings, reservationData);
            } catch (e) {
                log.error({ title: 'Christmas quantity Suitelet failed', details: e });
                fatal = e.message || String(e);
            }
            const allRows = Array.from(rows.values());
            allRows.sort((a, b) => (a.name || a.id).localeCompare(b.name || b.id) || a.id.localeCompare(b.id) || a.channel.localeCompare(b.channel));
            if (allRows.length > CONFIG.MAX_DISPLAY_ROWS) fatal = 'The full report exceeds CONFIG.MAX_DISPLAY_ROWS. Increase the configured limit after reviewing account capacity.';
            context.response.setHeader({ name: 'Content-Type', value: 'text/html; charset=UTF-8' });
            context.response.write(renderPage(allRows, allRows, '', '', warnings, fatal, inventory));
        }
        return { onRequest };
    });
