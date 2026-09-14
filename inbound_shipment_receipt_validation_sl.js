/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/search', 'N/record', 'N/runtime', 'N/url', 'N/file', 'N/log'],
(search, record, runtime, url, file, log) => {
    const TITLE = 'Inbound Shipment Receipt Validation';
    const PARAM = 'custscript_ibs_validation_search';
    const EPSILON = 0.00000001;
    const HEADER_FIELDS = ['internalid', 'shipmentnumber', 'custrecord157', 'custrecord158',
        'expectedshippingdate', 'custrecord_port_eta', 'memo', 'custrecord_mi_container_type', 'custrecord_conatiner_images'];
    const text = value => value == null ? '' : String(value);
    const join = column => text(column.join).toLowerCase();
    const esc = value => text(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    const yes = value => value === true || value === 'T';
    const number = value => Number(value || 0);

    function onRequest(context) {
        const action = context.request.parameters.action || '';
        try {
            if (!action) {
                context.response.write(buildPage());
                return;
            }
            let result;
            if (action === 'list') result = readSearch(context.request.parameters);
            else if (action === 'detail') result = getDetail(context.request.parameters);
            else if (action === 'save' && context.request.method === 'POST') {
                result = saveDetails(JSON.parse(context.request.body || '{}'));
            } else throw Error('Invalid request.');
            writeJson(context, {ok: true, data: result});
        } catch (error) {
            log.error({title: 'IBS validation: ' + (action || 'page'), details: error});
            if (action) writeJson(context, {ok: false, error: error.message || text(error)});
            else context.response.write('<h2>' + esc(TITLE) + '</h2><p>' + esc(error.message) + '</p>');
        }
    }

    function writeJson(context, value) {
        context.response.setHeader({name: 'Content-Type', value: 'application/json; charset=UTF-8'});
        context.response.setHeader({name: 'Cache-Control', value: 'no-store'});
        context.response.write(JSON.stringify(value));
    }

    // Load the configured search; keep its existing filters and column labels.
    function readSearch(filters = {}) {
        const searchId = runtime.getCurrentScript().getParameter({name: PARAM});
        if (!searchId) throw Error('Set the deployment parameter ' + PARAM + ' to your inbound shipment saved search ID.');
        const saved = search.load({id: searchId});
        if (text(saved.searchType).toLowerCase() !== 'inboundshipment') throw Error('The configured search must be an Inbound Shipment search.');
        if (saved.columns.some(c => c.summary)) throw Error('Use a detail saved search without summary/group columns.');
        const columns = saved.columns;
        const baseId = columns.find(c => !join(c) && c.name === 'internalid');
        const itemColumn = columns.find(c => !join(c) && c.name === 'item');
        if (!baseId || !itemColumn) throw Error('The search needs Internal Id and Item columns.');
        let lineColumn = columns.find(c => join(c) === 'inboundshipmentitem' && ['internalid', 'id'].includes(c.name));
        if (!lineColumn) lineColumn = search.createColumn({name: 'internalid', join: 'inboundShipmentItem', label: 'IBS Line ID'});
        saved.columns = columns.includes(lineColumn) ? columns : columns.concat(lineColumn);
        saved.columns = saved.columns.concat(search.createColumn({name: 'internalid', sort: search.Sort.ASC}),
            search.createColumn({name: lineColumn.name, join: lineColumn.join, sort: search.Sort.ASC}));
        const extra = [];
        [['ibs', 'shipmentnumber'], ['container', 'custrecord157'], ['seal', 'custrecord158']].forEach(([key, field]) => {
            if (text(filters[key]).trim()) extra.push(search.createFilter({name: 'formulatext', formula: '{' + field + '}', operator: search.Operator.CONTAINS, values: text(filters[key]).trim()}));
        });
        if (filters.shipmentId) extra.push(search.createFilter({name: 'internalid', operator: search.Operator.ANYOF, values: validId(filters.shipmentId)}));
        saved.filters = saved.filters.concat(extra);
        const visible = columns.filter(c => c !== lineColumn);
        const meta = visible.map((c, i) => ({key: 'c' + i, name: c.name, join: join(c), label: c.label || c.name,
            section: join(c) === 'inventorydetail' ? 'inventory' : (!join(c) && HEADER_FIELDS.includes(c.name) ? 'header' : 'item')}));
        const groups = new Map();
        const images = {};
        const pages = saved.runPaged({pageSize: 1000});
        pages.pageRanges.forEach(page => {
            if (runtime.getCurrentScript().getRemainingUsage() < 120) throw Error('Too many results. Narrow the shipment, container, or seal filters and try again.');
            pages.fetch({index: page.index}).data.forEach(result => {
                const shipmentId = text(result.getValue(baseId));
                const lineId = text(result.getValue(lineColumn));
                if (!lineId) throw Error('The search did not return an IBS item line ID. Check the Inbound Shipment Item: Internal ID column.');
                if (!groups.has(shipmentId)) groups.set(shipmentId, {id: shipmentId, cells: {}, lines: new Map()});
                const shipment = groups.get(shipmentId);
                if (!shipment.lines.has(lineId)) shipment.lines.set(lineId, {id: lineId, itemId: text(result.getValue(itemColumn)), cells: {}, hasDetail: false});
                const line = shipment.lines.get(lineId);
                visible.forEach((column, i) => {
                    const m = meta[i];
                    const value = result.getValue(column);
                    let label;
                    try { label = result.getText(column); } catch (_) { label = ''; }
                    const cell = {value: text(value), text: text(label || value)};
                    if (m.section === 'inventory') {
                        if (m.name === 'internalid' && value) line.hasDetail = true;
                        return;
                    }
                    let type = '';
                    if (!m.join && ['internalid', 'shipmentnumber'].includes(m.name)) type = 'inboundshipment';
                    if (!m.join && m.name === 'item') type = 'inventoryitem';
                    if (!m.join && m.name === 'vendor') type = 'vendor';
                    if (m.join === 'itemreceipt') type = 'itemreceipt';
                    if (type) {
                        let id = value;
                        if (type === 'inboundshipment') id = shipmentId;
                        if (type === 'itemreceipt') {
                            const idColumn = columns.find(c => join(c) === 'itemreceipt' && c.name === 'internalid');
                            id = idColumn ? result.getValue(idColumn) : '';
                        }
                        if (id) cell.url = url.resolveRecord({recordType: type, recordId: id, isEditMode: false});
                    }
                    if (['custrecord_conatiner_images', 'custitem_atlas_item_image'].includes(m.name) && value) {
                        if (!(text(value) in images)) {
                            try { images[text(value)] = file.load({id: value}).url; }
                            catch (error) { images[text(value)] = ''; log.debug({title: 'Image not available', details: {fileId: value, message: error.message}}); }
                        }
                        cell.image = images[text(value)];
                    }
                    const cells = m.section === 'header' ? shipment.cells : line.cells;
                    if (!cells[m.key]) cells[m.key] = [];
                    if (!cells[m.key].some(c => c.value === cell.value)) cells[m.key].push(cell);
                });
            });
        });
        const shipments = Array.from(groups.values()).map(s => ({...s, lines: Array.from(s.lines.values())}));
        log.debug({title: 'IBS search loaded', details: {searchId, resultRows: pages.count, shipments: shipments.length}});
        return {columns: meta, shipments};
    }

    function validId(value) {
        if (!/^\d+$/.test(text(value))) throw Error('Invalid record or line ID.');
        return text(value);
    }

    function findLine(shipment, lineId) {
        for (let line = 0; line < shipment.getLineCount({sublistId: 'items'}); line++) {
            if (text(shipment.getSublistValue({sublistId: 'items', fieldId: 'id', line})) === text(lineId)) return line;
        }
        throw Error('The shipment item line no longer exists. Refresh the page.');
    }

    function isoDate(value) {
        if (!value) return '';
        if (!(value instanceof Date) || isNaN(value.getTime())) throw Error('Unexpected inventory expiration date.');
        return value.getFullYear() + '-' + String(value.getMonth() + 1).padStart(2, '0') + '-' + String(value.getDate()).padStart(2, '0');
    }

    function parseDate(value) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw Error('Enter a valid expiration date.');
        const parts = value.split('-').map(Number);
        const date = new Date(parts[0], parts[1] - 1, parts[2]);
        if (isoDate(date) !== value) throw Error('Enter a valid expiration date.');
        return date;
    }

    function readAssignments(shipment, line) {
        if (!shipment.hasSublistSubrecord({sublistId: 'items', fieldId: 'inventorydetail', line})) return {id: '', rows: []};
        const detail = shipment.getSublistSubrecord({sublistId: 'items', fieldId: 'inventorydetail', line});
        const fields = detail.getSublistFields({sublistId: 'inventoryassignment'});
        const rows = [];
        for (let index = 0; index < detail.getLineCount({sublistId: 'inventoryassignment'}); index++) {
            const value = field => fields.includes(field) ? detail.getSublistValue({sublistId: 'inventoryassignment', fieldId: field, line: index}) : '';
            const label = field => {
                if (!fields.includes(field)) return '';
                return text(detail.getSublistText({sublistId: 'inventoryassignment', fieldId: field, line: index}));
            };
            rows.push({number: text(value('receiptinventorynumber') || label('issueinventorynumber')),
                bin: text(value('binnumber')), status: text(value('inventorystatus')),
                expiry: isoDate(value('expirationdate')), quantity: number(value('quantity'))});
        }
        return {id: text(shipment.getSublistValue({sublistId: 'items', fieldId: 'inventorydetail', line})), rows};
    }

    function lineState(shipment, line) {
        const value = fieldId => shipment.getSublistValue({sublistId: 'items', fieldId, line});
        return {expected: number(value('quantityexpected')), received: number(value('quantityreceived')),
            locationId: text(value('receivinglocation')), unit: text(value('unit')),
            purchaseOrder: text(value('purchaseorder')), shipmentItem: text(value('shipmentitem')),
            inventory: readAssignments(shipment, line)};
    }

    function options(type, filters, nameField) {
        const values = [];
        const list = search.create({type, filters, columns: ['internalid', nameField]}).runPaged({pageSize: 1000});
        list.pageRanges.forEach(p => list.fetch({index: p.index}).data.forEach(r => values.push({id: text(r.getValue('internalid')), name: text(r.getValue(nameField))})));
        return values;
    }

    function itemRules(itemId) {
        const item = search.lookupFields({type: search.Type.ITEM, id: itemId, columns: ['islotitem', 'isserialitem', 'usebins']});
        return {lot: yes(item.islotitem), serial: yes(item.isserialitem), bins: yes(item.usebins),
            statuses: runtime.isFeatureInEffect({feature: 'INVENTORYSTATUS'})};
    }

    function unitInfo(itemId, unitId) {
        if (!runtime.isFeatureInEffect({feature: 'MULTIPLEUNITS'}) || !unitId) return {rate: 1, name: ''};
        const item = search.lookupFields({type: search.Type.ITEM, id: itemId, columns: ['unitstype']});
        const typeId = item.unitstype && item.unitstype[0] && item.unitstype[0].value;
        if (!typeId) throw Error('Cannot resolve the item units type.');
        const units = record.load({type: 'unitstype', id: typeId});
        let rate = 0, name = '';
        for (let line = 0; line < units.getLineCount({sublistId: 'uom'}); line++) {
            const get = fieldId => units.getSublistValue({sublistId: 'uom', fieldId, line});
            if (text(get('internalid')) === unitId) rate = Number(get('conversionrate'));
            if (yes(get('baseunit'))) name = text(get('unitname'));
        }
        if (!Number.isFinite(rate) || rate <= 0) throw Error('Cannot resolve the shipment line unit conversion.');
        return {rate, name};
    }

    function getDetail(params, loadedShipment, searchData) {
        const shipmentId = validId(params.shipmentId);
        const lineId = validId(params.lineId);
        const data = searchData || readSearch({shipmentId});
        const searchLine = data.shipments.find(s => s.id === shipmentId)?.lines.find(l => l.id === lineId);
        if (!searchLine) throw Error('This shipment line is no longer included in the configured search.');
        const shipment = loadedShipment || record.load({type: 'inboundshipment', id: shipmentId});
        const index = findLine(shipment, lineId);
        const state = lineState(shipment, index);
        const rules = itemRules(searchLine.itemId);
        const units = unitInfo(searchLine.itemId, state.unit);
        const expected = state.expected * units.rate, received = state.received * units.rate;
        const bins = rules.bins && state.locationId ? options('bin', [['location','anyof',state.locationId], 'AND', ['inactive','is','F']], 'binnumber') : [];
        const statuses = rules.statuses ? options('inventorystatus', [['isinactive','is','F']], 'name') : [];
        return {shipmentId, lineId, ...state, expected, received, units, rules, bins, statuses, snapshot: JSON.stringify(state),
            location: text(shipment.getSublistText({sublistId: 'items', fieldId: 'receivinglocation', line: index})),
            item: searchLine.cells[data.columns.find(c => !c.join && c.name === 'item').key][0].text,
            max: Math.max(0, expected - received), columns: data.columns.filter(c => c.section === 'inventory')};
    }

    // Recheck current quantities and assignments before saving the IBS record.
    function validateRows(rows, detail) {
        if (!Array.isArray(rows)) throw Error('Invalid inventory detail rows.');
        let total = 0;
        const serials = new Set();
        rows.forEach((row, index) => {
            const prefix = 'Row ' + (index + 1) + ': ';
            row.quantity = Number(row.quantity);
            row.number = text(row.number).trim();
            if (!Number.isFinite(row.quantity) || row.quantity <= 0) throw Error(prefix + 'quantity must be greater than zero.');
            if ((detail.rules.lot || detail.rules.serial) && !row.number) throw Error(prefix + 'enter a lot/serial number.');
            if (detail.rules.serial) {
                if (row.quantity !== 1) throw Error(prefix + 'serial quantity must be 1.');
                if (serials.has(row.number.toLowerCase())) throw Error(prefix + 'duplicate serial number.');
                serials.add(row.number.toLowerCase());
            }
            if (row.bin && !detail.bins.some(b => b.id === text(row.bin))) throw Error(prefix + 'bin must be active and belong to the receiving location.');
            if (row.status && !detail.statuses.some(s => s.id === text(row.status))) throw Error(prefix + 'select an active inventory status.');
            if (detail.rules.bins && !row.bin) throw Error(prefix + 'select a bin.');
            if (detail.rules.statuses && !row.status) throw Error(prefix + 'select a status.');
            if (row.expiry && !detail.rules.lot) throw Error(prefix + 'expiration dates apply to lot items.');
            if (row.expiry) parseDate(row.expiry);
            total += row.quantity;
        });
        if (total - detail.max > EPSILON) throw Error('Inventory detail total ' + total + ' exceeds Qty Expected minus Qty Received (' + detail.max + ').');
        return total;
    }

    function saveDetails(payload) {
        const shipmentId = validId(payload.shipmentId);
        if (!Array.isArray(payload.lines) || !payload.lines.length) throw Error('No changed item lines to submit.');
        const shipment = record.load({type: 'inboundshipment', id: shipmentId, isDynamic: false});
        const searchData = readSearch({shipmentId});
        const changed = new Set();
        for (const change of payload.lines) {
            if (runtime.getCurrentScript().getRemainingUsage() < 180) throw Error('Too many edited lines in one submission. Submit fewer lines at a time. No changes were saved for this shipment.');
            const lineId = validId(change.lineId);
            if (changed.has(lineId)) throw Error('Duplicate item line in submission.');
            changed.add(lineId);
            const detail = getDetail({shipmentId, lineId}, shipment, searchData);
            const index = findLine(shipment, lineId);
            if (detail.snapshot !== change.snapshot) {
                throw Error('Shipment quantities or inventory details changed for ' + detail.item + '. Reopen the popup before submitting.');
            }
            const total = validateRows(change.rows, detail);
            const subrecord = shipment.getSublistSubrecord({sublistId: 'items', fieldId: 'inventorydetail', line: index});
            for (let i = subrecord.getLineCount({sublistId: 'inventoryassignment'}) - 1; i >= 0; i--) {
                subrecord.removeLine({sublistId: 'inventoryassignment', line: i});
            }
            change.rows.forEach((row, line) => {
                const set = (fieldId, value) => subrecord.setSublistValue({sublistId: 'inventoryassignment', fieldId, line, value});
                if (detail.rules.lot || detail.rules.serial) set('receiptinventorynumber', row.number);
                if (row.expiry) set('expirationdate', parseDate(row.expiry));
                if (detail.rules.bins && row.bin) set('binnumber', Number(row.bin));
                if (detail.rules.statuses && row.status) set('inventorystatus', Number(row.status));
                set('quantity', row.quantity);
            });
            log.debug({title: 'IBS line validated', details: {shipmentId, lineId, rows: change.rows.length, total, maximum: detail.max}});
        }
        const id = shipment.save({enableSourcing: true, ignoreMandatoryFields: false});
        log.audit({title: 'IBS inventory details saved', details: {shipmentId: id, lines: Array.from(changed), userId: runtime.getCurrentUser().id}});
        return {id, lines: changed.size};
    }

    function buildPage() {
        const endpoint = url.resolveScript({scriptId: runtime.getCurrentScript().id, deploymentId: runtime.getCurrentScript().deploymentId});
        return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + TITLE + '</title><style>' + styles() + '</style></head><body>' +
            '<main><header><div><h1>' + TITLE + '</h1><p>Review shipments and validate inventory details</p></div><div><button id="refresh">Refresh</button> <button id="submit" class="primary" disabled>Submit</button></div></header>' +
            '<section class="metrics"><div><b id="shipCount">0</b>Shipments</div><div><b id="lineCount">0</b>Item Lines</div><div><b id="editCount">0</b>Changed Lines</div></section>' +
            '<form id="filters"><label>Shipment Number<input name="ibs" placeholder="Contains shipment number"></label><label>Container Number<input name="container" placeholder="Contains container number"></label><label>Seal Number<input name="seal" placeholder="Contains seal number"></label><button class="primary">Apply Filters</button><button type="button" id="clear">Clear</button></form>' +
            '<div id="message" role="status"></div><div class="table-wrap"><table id="shipments"></table></div><footer>Open + to view item lines. Click Inventory Detail to view or edit assignments.</footer></main>' +
            '<div id="modal" class="modal" role="dialog" aria-modal="true" aria-label="Inventory Detail" hidden><div class="dialog"><div class="dialog-head"><strong>Inventory Detail</strong><button id="close">×</button></div><div id="detailBody" class="dialog-body"></div><div class="dialog-actions"><button id="cancel">Cancel</button><button id="ok" class="primary">OK</button></div></div></div>' +
            '<div id="imageModal" class="modal" role="dialog" aria-modal="true" aria-label="Image preview" hidden><div class="image-dialog"><button id="closeImage">Close</button><img id="largeImage" alt="Full size image"></div></div>' +
            '<script>(' + client.toString() + ')(' + JSON.stringify(endpoint).replace(/</g, '\\u003c') + ');</script></body></html>';
    }

    function styles() {
        return `*{box-sizing:border-box}body{margin:0;background:#f3f6fa;color:#24364b;font:13px Arial,sans-serif}main{margin:20px;border:1px solid #d7dde7;border-radius:8px;overflow:hidden;background:white}header{background:#122d4a;color:white;padding:22px;display:flex;justify-content:space-between;align-items:center;gap:20px}h1{font-size:22px;margin:0}header p{margin:7px 0 0;color:#c4d2e1}button{cursor:pointer;background:white;color:#24364b;border:1px solid #bdc7d8;border-radius:5px;padding:8px 13px;font-weight:bold}button.primary{background:#1664c0;border-color:#1664c0;color:white}button:disabled{opacity:.5;cursor:default}.metrics{display:flex;gap:14px;background:#f8fafc;padding:18px}.metrics>div{background:white;border:1px solid #dce3ed;border-radius:7px;padding:14px 22px;min-width:160px;color:#64748b}.metrics b{display:block;font-size:25px;color:#163b63;margin-bottom:5px}form{display:flex;gap:12px;padding:16px;align-items:end;border-bottom:1px solid #d7dde7;flex-wrap:wrap}label{display:block;font-size:11px;font-weight:bold}input,select{display:block;margin-top:5px;width:100%;border:1px solid #bdc7d8;border-radius:5px;padding:7px;background:white;color:#24364b}form input{width:230px}#message{padding:12px 16px;white-space:pre-wrap}#message:empty{display:none}.error{color:#a52727;background:#fff0ef}.success{color:#17603c;background:#edf9f2}.table-wrap{overflow:auto;max-height:65vh}table{border-collapse:separate;border-spacing:0;width:100%;font-size:12px}th{background:#e8eef7;color:#24364b;border-bottom:1px solid #cad4e3;border-right:1px solid #d8e0ea;padding:10px;text-align:left;white-space:nowrap}#shipments>thead th{position:sticky;top:0;z-index:1}td{border-bottom:1px solid #e6ebf2;border-right:1px solid #edf1f6;padding:9px;vertical-align:middle;min-width:90px}tr:hover>td{background:#f8fbff}td:first-child{min-width:40px}a{color:#165ba7;text-decoration:none;font-weight:bold}a:hover{text-decoration:underline}.expanded>td{padding:14px;background:#f9fbfd}.item-scroll{overflow:auto;max-width:calc(100vw - 100px)}.item-scroll table{min-width:1120px}.expander{padding:2px;width:24px;height:24px}.thumb{width:45px;height:45px;object-fit:contain;cursor:zoom-in}.detail-button{color:#1664c0;font-size:19px;padding:3px 8px}.filled{color:#168052}.dirty{background:#fff4d5!important}footer{padding:13px 16px;color:#667085}.modal{position:fixed;inset:0;z-index:10;display:flex;align-items:center;justify-content:center;background:rgba(15,23,42,.38)}[hidden]{display:none!important}.dialog{width:min(1100px,calc(100vw - 32px));max-height:calc(100vh - 44px);overflow:auto;background:white;border:1px solid #c9d4e4;border-radius:8px;box-shadow:0 24px 70px rgba(15,23,42,.24)}.dialog-head{display:flex;justify-content:space-between;align-items:center;padding:12px 14px;background:#e8eef7;border-bottom:1px solid #cad4e3}.dialog-head strong{font-size:15px}.dialog-body{padding:16px}.dialog-actions{display:flex;justify-content:flex-end;gap:8px;padding:12px 14px;border-top:1px solid #d7dde7}.detail-summary{display:flex;gap:25px;flex-wrap:wrap;margin-bottom:15px}.detail-summary b{display:block;margin-top:4px}.inventory-wrap{overflow:auto}.inventory-wrap input,.inventory-wrap select{min-width:110px}.inventory-wrap input[type=number]{width:95px;min-width:95px}.image-dialog{background:white;padding:15px;border-radius:8px;max-width:94vw}.image-dialog img{display:block;max-width:90vw;max-height:80vh;margin-top:10px}.hint{color:#667085;padding:10px 0}.empty{text-align:center;padding:40px;color:#64748b}@media(max-width:700px){main{margin:8px}header{align-items:flex-start;flex-direction:column}.metrics{gap:6px}.metrics>div{min-width:0;padding:12px;flex:1}h1{font-size:19px}}`;
    }

    function client(endpoint) {
        const $ = id => document.getElementById(id);
        const escape = value => String(value == null ? '' : value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
        let data = {columns: [], shipments: []}, expanded = new Set(), edits = {}, active = null, busy = false;
        const key = (shipmentId, lineId) => shipmentId + ':' + lineId;
        const message = (value, error) => { $('message').textContent = value; $('message').className = error ? 'error' : 'success'; };
        async function request(action, params, payload) {
            const address = new URL(endpoint, location.origin);
            address.searchParams.set('action', action);
            Object.entries(params || {}).forEach(([k,v]) => address.searchParams.set(k, v));
            const response = await fetch(address, payload ? {method:'POST', credentials:'same-origin', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)} : {credentials:'same-origin', cache:'no-store'});
            const result = await response.json();
            if (!result.ok) throw Error(result.error || 'Request failed.');
            return result.data;
        }
        function setBusy(value) {
            busy = value;
            $('submit').disabled = value || !Object.keys(edits).length;
            $('refresh').disabled = value;
            $('filters').querySelectorAll('button,input').forEach(e => e.disabled = value);
        }
        function cell(values) {
            return (values || []).map(c => {
                if (c.image) return '<img class="thumb" src="' + escape(c.image) + '" data-image="' + escape(c.image) + '" alt="' + escape(c.text) + '">';
                return c.url ? '<a target="_blank" rel="noopener" href="' + escape(c.url) + '">' + escape(c.text) + '</a>' : escape(c.text);
            }).join('<br>');
        }
        function render() {
            const headers = data.columns.filter(c => c.section === 'header');
            const items = data.columns.filter(c => c.section === 'item');
            let html = '<thead><tr><th></th>' + headers.map(c => '<th>' + escape(c.label) + '</th>').join('') + '</tr></thead><tbody>';
            data.shipments.forEach(s => {
                html += '<tr><td><button class="expander" aria-expanded="' + expanded.has(s.id) + '" data-expand="' + s.id + '">' + (expanded.has(s.id) ? '−' : '+') + '</button></td>' + headers.map(c => '<td>' + cell(s.cells[c.key]) + '</td>').join('') + '</tr>';
                if (expanded.has(s.id)) {
                    html += '<tr class="expanded"><td colspan="' + (headers.length + 1) + '"><div class="item-scroll"><table><thead><tr>' + items.map(c => '<th>' + escape(c.label) + '</th>').join('') + '<th>Inventory Detail</th></tr></thead><tbody>';
                    s.lines.forEach(l => {
                        const draft = edits[key(s.id,l.id)];
                        html += '<tr>' + items.map(c => '<td>' + cell(l.cells[c.key]) + '</td>').join('') + '<td class="' + (draft ? 'dirty' : '') + '"><button class="detail-button ' + ((draft ? draft.rows.length : l.hasDetail) ? 'filled' : '') + '" title="View / Edit Inventory Detail" aria-label="View / Edit Inventory Detail" data-detail="' + s.id + ':' + l.id + '">' + ((draft ? draft.rows.length : l.hasDetail) ? '▤✓' : '▤+') + '</button></td></tr>';
                    });
                    html += '</tbody></table></div></td></tr>';
                }
            });
            html += '</tbody>';
            if (!data.shipments.length) html += '<tbody><tr><td class="empty" colspan="' + (headers.length+1) + '">No matching shipments.</td></tr></tbody>';
            $('shipments').innerHTML = html;
            $('shipCount').textContent = data.shipments.length;
            $('lineCount').textContent = data.shipments.reduce((sum,s) => sum+s.lines.length,0);
            $('editCount').textContent = Object.keys(edits).length;
            setBusy(busy);
        }
        async function load(discard) {
            if (busy) return;
            if (discard && Object.keys(edits).length && !confirm('Discard unsaved inventory changes and refresh?')) return;
            if (discard) edits = {};
            setBusy(true); message('Loading shipments…');
            try { data = await request('list', Object.fromEntries(new FormData($('filters')))); message(''); render(); }
            catch(error) { message(error.message,true); }
            finally { setBusy(false); }
        }
        function options(list, value) {
            let html = '<option value="">Select</option>';
            if (value && !list.some(v => v.id === value)) html += '<option selected value="' + escape(value) + '">Unavailable (' + escape(value) + ')</option>';
            return html + list.map(v => '<option value="' + escape(v.id) + '"' + (v.id === value ? ' selected' : '') + '>' + escape(v.name) + '</option>').join('');
        }
        function inventoryCell(column, row, index) {
            const attr = ' data-row="' + index + '" ';
            switch(column.name) {
                case 'internalid': return escape(active.inventory.id);
                case 'item': return escape(active.item);
                case 'location': return escape(active.location);
                case 'inventorynumber': return active.rules.lot || active.rules.serial ? '<input' + attr + 'data-field="number" value="' + escape(row.number) + '">' : '—';
                case 'binnumber': return active.rules.bins ? '<select' + attr + 'data-field="bin">' + options(active.bins,row.bin) + '</select>' : '—';
                case 'status': return active.rules.statuses ? '<select' + attr + 'data-field="status">' + options(active.statuses,row.status) + '</select>' : '—';
                case 'expirationdate': return active.rules.lot ? '<input type="date"' + attr + 'data-field="expiry" value="' + escape(row.expiry) + '">' : '—';
                case 'quantity': return '<input type="number" min="0" step="any"' + attr + 'data-field="quantity" value="' + escape(row.quantity) + '">';
                default: return '—';
            }
        }
        function showDetail() {
            const required = [['inventorynumber','Number'],['binnumber','Bin Number'],['expirationdate','Expiration Date'],['quantity','Quantity'],['status','Status']];
            const columns = active.columns.slice();
            required.forEach(([name,label]) => { if (!columns.some(c => c.name === name)) columns.push({name,label}); });
            $('detailBody').innerHTML = '<div class="hint">Quantities in ' + escape(active.units.name || 'base units') + '</div><div class="detail-summary"><div>Item<b>' + escape(active.item) + '</b></div><div>Qty Expected<b>' + active.expected + '</b></div><div>Qty Received<b>' + active.received + '</b></div><div>Maximum Quantity<b>' + active.max + '</b></div><div>Total Qty<b id="total"></b></div></div>' +
                '<div class="inventory-wrap"><table><thead><tr>' + columns.map(c => '<th>' + escape(c.label) + '</th>').join('') + '<th></th></tr></thead><tbody>' + active.rows.map((row,i) => '<tr>' + columns.map(c => '<td>' + inventoryCell(c,row,i) + '</td>').join('') + '<td><button data-remove="' + i + '">Remove</button></td></tr>').join('') + '</tbody></table></div><p><button id="addRow">Add Row</button></p><div class="hint">Total Qty must not exceed Qty Expected − Qty Received. OK stages changes; Submit saves the shipment.</div><div id="detailError" class="error" role="alert"></div>';
            updateTotal();
        }
        function updateTotal() { $('total').textContent = Math.round(active.rows.reduce((sum,r) => sum + (Number(r.quantity)||0),0)*1e8)/1e8; }
        async function openDetail(shipmentId,lineId) {
            if (busy) return;
            setBusy(true);
            try {
                const latest = await request('detail', {shipmentId,lineId});
                const draft = edits[key(shipmentId,lineId)];
                if (draft && draft.snapshot !== latest.snapshot) {
                    delete edits[key(shipmentId,lineId)];
                    message('This line changed in NetSuite. Its previous draft was discarded; current inventory details are shown.',true);
                }
                active = {...latest, rows: JSON.parse(JSON.stringify(draft && draft.snapshot === latest.snapshot ? draft.rows : latest.inventory.rows))};
                showDetail(); $('modal').hidden = false; $('close').focus();
            } catch(error) { message(error.message,true); }
            finally { setBusy(false); render(); }
        }
        function closeDetail() { $('modal').hidden = true; active = null; }
        function stageDetail() {
            let total = 0; const serials = new Set();
            for (const row of active.rows) {
                const q = Number(row.quantity);
                let error = '';
                if (!Number.isFinite(q) || q <= 0) error = 'Enter a quantity greater than zero on every row.';
                if ((active.rules.lot || active.rules.serial) && !row.number.trim()) error = 'Enter each lot/serial number.';
                if (active.rules.bins && !row.bin) error = 'Select each bin.';
                if (active.rules.statuses && !row.status) error = 'Select each status.';
                if (active.rules.serial && (q !== 1 || serials.has(row.number.trim().toLowerCase()))) error = 'Use unique serial numbers with quantity 1.';
                serials.add(row.number.trim().toLowerCase()); total += q;
                if (error) { $('detailError').textContent = error; return; }
            }
            if (total-active.max > 0.00000001) { $('detailError').textContent = 'Total inventory quantity cannot exceed ' + active.max + '.'; return; }
            const k = key(active.shipmentId,active.lineId);
            if (JSON.stringify(active.rows) === JSON.stringify(active.inventory.rows)) delete edits[k];
            else edits[k] = {shipmentId:active.shipmentId,lineId:active.lineId,snapshot:active.snapshot,rows:active.rows};
            closeDetail(); render();
        }
        async function submit() {
            if (busy || !Object.keys(edits).length) return;
            setBusy(true); message('Validating and saving inventory details…');
            const groups = {};
            Object.values(edits).forEach(e => (groups[e.shipmentId] ||= []).push(e));
            let saved = 0;
            try {
                for (const [shipmentId,lines] of Object.entries(groups)) {
                    await request('save', {}, {shipmentId,lines});
                    lines.forEach(l => delete edits[key(shipmentId,l.lineId)]);
                    saved++;
                }
                data = await request('list', Object.fromEntries(new FormData($('filters'))));
                message('Inventory details saved for ' + saved + ' shipment(s).');
            } catch(error) { message((saved ? saved + ' shipment(s) saved. ' : '') + error.message + ' Remaining drafts are retained. Reopen the affected popup if the record changed.',true); }
            finally { setBusy(false); render(); }
        }
        $('filters').onsubmit = event => { event.preventDefault(); load(true); };
        $('refresh').onclick = () => load(true);
        $('clear').onclick = () => { if (Object.keys(edits).length && !confirm('Discard unsaved inventory changes and clear filters?')) return; edits = {}; $('filters').reset(); load(false); };
        $('submit').onclick = submit;
        $('shipments').onclick = event => {
            const button = event.target.closest('button');
            if (button && !busy) {
                if (button.dataset.expand) { const id = button.dataset.expand; expanded.has(id) ? expanded.delete(id) : expanded.add(id); render(); }
                if (button.dataset.detail) openDetail(...button.dataset.detail.split(':'));
            }
            if (event.target.dataset.image) { $('largeImage').src = event.target.dataset.image; $('imageModal').hidden = false; $('closeImage').focus(); }
        };
        $('detailBody').oninput = event => {
            const field = event.target.dataset.field;
            if (field) { active.rows[Number(event.target.dataset.row)][field] = event.target.value; updateTotal(); }
        };
        $('detailBody').onclick = event => {
            if (event.target.id === 'addRow') { active.rows.push({number:'',bin:'',status:'',expiry:'',quantity:active.rules.serial ? 1 : ''}); showDetail(); }
            if (event.target.dataset.remove !== undefined) { active.rows.splice(Number(event.target.dataset.remove),1); showDetail(); }
        };
        $('close').onclick = closeDetail; $('cancel').onclick = closeDetail; $('ok').onclick = stageDetail;
        $('closeImage').onclick = () => $('imageModal').hidden = true;
        document.onkeydown = event => {
            if (event.key === 'Escape') { if (!$('imageModal').hidden) $('imageModal').hidden = true; else if (!$('modal').hidden) closeDetail(); }
        };
        window.onbeforeunload = event => { if (Object.keys(edits).length) { event.preventDefault(); event.returnValue = ''; } };
        load(false);
    }

    return {onRequest};
});
