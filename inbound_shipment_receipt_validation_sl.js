/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/search', 'N/record', 'N/runtime', 'N/url', 'N/file', 'N/log', 'N/format'],
(search, record, runtime, url, file, log, format) => {
    const TITLE = 'Inbound Shipment Receipt Validation';
    const QC_FIELD = 'custrecord_mi_qc_status';
    const PARAM = 'custscript_ibs_validation_search';
    let lineMaps = new WeakMap();
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
        const started = Date.now();
        lineMaps = new WeakMap();
        try {
            if (!action) {
                context.response.write(buildPage());
                return;
            }
            let result;
            if (action === 'list') result = readSearch(context.request.parameters);
            else if (action === 'save' && context.request.method === 'POST') {
                result = saveDetails(JSON.parse(context.request.body || '{}'));
            } else throw Error('Invalid request.');
            writeJson(context, {ok: true, data: result});
        } catch (error) {
            log.error({title: 'IBS validation: ' + (action || 'page'), details: error});
            if (action) writeJson(context, {ok: false, error: error.message || text(error)});
            else context.response.write('<h2>' + esc(TITLE) + '</h2><p>' + esc(error.message) + '</p>');
        } finally {
            log.debug({title: 'IBS request timing', details: {action: action || 'page', milliseconds: Date.now() - started, remainingUsage: runtime.getCurrentScript().getRemainingUsage()}});
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
        const columns = saved.columns.filter(c => !(join(c) === 'inboundshipmentitem' && ['internalid', 'id'].includes(c.name)));
        const baseId = columns.find(c => !join(c) && c.name === 'internalid');
        const itemColumn = columns.find(c => !join(c) && c.name === 'item');
        if (!baseId || !itemColumn) throw Error('The search needs Internal Id and Item columns.');
        const lineColumn = itemColumn;
        const locationColumn = columns.find(c => !join(c) && c.name === 'receivinglocation') || search.createColumn({name:'receivinglocation'});
        const unitColumn = columns.find(c => !join(c) && c.name === 'unit') || search.createColumn({name:'unit'});
        saved.columns = columns.concat([locationColumn, unitColumn].filter(c => !columns.includes(c)));
        const extra = [];
        [['ibs', 'shipmentnumber'], ['container', 'custrecord157'], ['seal', 'custrecord158']].forEach(([key, field]) => {
            if (text(filters[key]).trim()) extra.push(search.createFilter({name: 'formulatext', formula: '{' + field + '}', operator: search.Operator.CONTAINS, values: text(filters[key]).trim()}));
        });
        if (filters.shipmentId) extra.push(search.createFilter({name: 'internalid', operator: search.Operator.ANYOF, values: validId(filters.shipmentId)}));
        saved.filters = saved.filters.concat(extra);
        const visible = columns.filter(c => !((!join(c) || join(c) === 'itemreceipt') && c.name === 'internalid'));
        const meta = visible.map((c, i) => ({key: 'c' + i, name: c.name, join: join(c), label: c.label || c.name,
            section: join(c) === 'inventorydetail' ? 'inventory' : (!join(c) && HEADER_FIELDS.includes(c.name) ? 'header' : 'item')}));
        const groups = new Map();
        const inventoryStates = new Map();
        let rowNumber = 0;
        const images = {};
        const links = {};
        log.debug({title: 'IBS item identity column', details: {name: lineColumn.name, join: lineColumn.join}});
        const pages = saved.runPaged({pageSize: 1000});
        pages.pageRanges.forEach(page => {
            if (runtime.getCurrentScript().getRemainingUsage() < 120) throw Error('The configured saved search returns too many results. Narrow its criteria and refresh.');
            pages.fetch({index: page.index}).data.forEach(result => {
                const shipmentId = text(result.getValue(baseId));
                const itemId = text(result.getValue(itemColumn));
                const lineId = String(++rowNumber);
                if (!itemId) throw Error('The search returned an empty Item value.');
                if (!groups.has(shipmentId)) groups.set(shipmentId, {id: shipmentId, cells: {}, lines: new Map()});
                const shipment = groups.get(shipmentId);
                shipment.lines.set(lineId, {id: lineId, itemId, cells: {}, hasDetail: false});
                const line = shipment.lines.get(lineId);
                const get = (name, joined) => {
                    const column = columns.find(c => c.name === name && join(c) === (joined || ''));
                    return column ? result.getValue(column) : '';
                };
                line.poId = text(get('purchaseorder'));
                const qcColumn = columns.find(c => c.name === QC_FIELD);
                const issueColumn = columns.find(c => c.name === 'custrecord_mi_open_issue');
                line.qcOriginal = text(qcColumn ? result.getValue(qcColumn) : '');
                const issue = text(issueColumn ? result.getText(issueColumn) || result.getValue(issueColumn) : '').trim();
                line.qcInitial = qcColumn ? qcDefault(line.qcOriginal, issue) : '';
                line.qcEnabled = !!qcColumn;
                const inventoryKey = shipmentId + ':' + itemId + ':' + text(get('internalid','inventorydetail'));
                const inventoryKey = shipmentId + ':' + line.poId + ':' + itemId + ':' + text(get('internalid','inventorydetail'));
                if (!inventoryStates.has(inventoryKey)) inventoryStates.set(inventoryKey, {expected:number(get('quantityexpected')), received:number(get('quantityreceived')),
                    locationId:text(result.getValue(locationColumn)), location:text(result.getText(locationColumn) || ''),
                    unit:text(result.getValue(unitColumn)), inventory:{id:text(get('internalid','inventorydetail')),rows:[]}});
                line.stored = inventoryStates.get(inventoryKey);
                const qty = get('quantity','inventorydetail');
                if (qty !== '' && qty != null && Number(qty) !== 0) {
                    const lotColumn = columns.find(c => c.name === 'inventorynumber' && join(c) === 'inventorydetail');
                    const expiry = get('expirationdate','inventorydetail');
                    const row = {number:text(lotColumn ? result.getText(lotColumn) || result.getValue(lotColumn) : ''),
                        bin:text(get('binnumber','inventorydetail')), status:text(get('status','inventorydetail')),
                        expiry:expiry ? isoDate(format.parse({value:text(expiry),type:format.Type.DATE})) : '', quantity:Number(qty)};
                    if (!line.stored.inventory.rows.some(r => JSON.stringify(r) === JSON.stringify(row))) line.stored.inventory.rows.push(row);
                }
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
                    const cells = m.section === 'header' ? shipment.cells : line.cells;
                    if (!cells[m.key]) cells[m.key] = [];
                    if (cells[m.key].some(c => c.value === cell.value)) return;
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
                        if (id) {
                            const linkKey = type + ':' + id;
                            if (!links[linkKey]) links[linkKey] = url.resolveRecord({recordType: type, recordId: id, isEditMode: false});
                            cell.url = links[linkKey];
                        }
                    }
                    if (['custrecord_conatiner_images', 'custitem_atlas_item_image'].includes(m.name)) {
                        const imageKey = text(value || label);
                        if (!(imageKey in images)) images[imageKey] = imageUrl(value, label);
                        cell.image = images[imageKey];
                        cell.isImage = true;
                    }
                    cells[m.key].push(cell);
                });
            });
        });
        const shipments = Array.from(groups.values()).map(s => ({...s, lines: Array.from(s.lines.values())}));
        preloadDetails(shipments, meta);
        log.debug({title: 'IBS search loaded', details: {searchId, resultRows: pages.count, shipments: shipments.length}});
        return {columns: meta, shipments};
    }

    function qcDefault(existing, issue) {
        return text(existing) || (text(issue).trim().toLowerCase() === 'label' ? '2' : text(issue).trim() ? '3' : '');
    }

    // Prepare popup data once, from the main search, before returning the page data.
    function preloadDetails(shipments, columns) {
        const rulesByItem = {}, unitsByItem = {}, binsByLocation = {};
        let statuses;
        const itemColumn = columns.find(c => !c.join && c.name === 'item');
        shipments.forEach(shipment => shipment.lines.forEach(line => {
            const state = line.stored;
            if (!rulesByItem[line.itemId]) rulesByItem[line.itemId] = itemRules(line.itemId);
            const rules = rulesByItem[line.itemId];
            const unitKey = line.itemId + ':' + state.unit;
            if (!unitsByItem[unitKey]) unitsByItem[unitKey] = unitInfo(line.itemId, state.unit);
            const units = unitsByItem[unitKey];
            if (rules.bins && state.locationId && !binsByLocation[state.locationId]) {
                binsByLocation[state.locationId] = options('bin', [['location','anyof',state.locationId], 'AND', ['inactive','is','F']], 'binnumber');
            }
            if (rules.statuses && !statuses) statuses = options('inventorystatus', [['isinactive','is','F']], 'name');
            const expected = state.expected * units.rate, received = state.received * units.rate;
            line.detail = {...state,shipmentId:shipment.id,lineId:line.id,itemId:line.itemId,expected,received,units,rules,
            line.detail = {...state,shipmentId:shipment.id,lineId:line.id,itemId:line.itemId,poId:line.poId,expected,received,units,rules,
                bins:binsByLocation[state.locationId] || [],statuses:rules.statuses ? statuses : [],
                item:line.cells[itemColumn.key][0].text,max:Math.max(0,expected-received),
                columns:columns.filter(c => c.section === 'inventory'),snapshot:snapshot(state)};
            delete line.stored;
        }));
    }

    function snapshot(state) {
        const rows = state.inventory.rows.map(r => JSON.stringify({number:text(r.number),bin:text(r.bin),status:text(r.status),expiry:text(r.expiry),quantity:Number(r.quantity)})).sort();
        return JSON.stringify({expected:state.expected,received:state.received,locationId:state.locationId,rows});
    }

    function imageUrl(value, label) {
        for (const candidate of [value, label]) {
            let source = text(candidate).trim().replace(/&amp;/gi, '&');
            if (!source) continue;
            const attribute = source.match(/(?:src|href)\s*=\s*["']([^"']+)["']/i);
            if (attribute) source = attribute[1];
            const mediaId = source.match(/[?&]id=(\d+)/i);
            const fileId = /^\d+$/.test(source) ? source : (mediaId ? mediaId[1] : '');
            if (fileId) {
                try { return file.load({id: fileId}).url; }
                catch (error) { log.debug({title: 'Image file lookup failed', details: {fileId, message: error.message}}); }
            }
            if (/^(https?:\/\/|\/(?!\/))/i.test(source) && !/[<>"'\r\n]/.test(source)) return source;
        }
        return '';
    }

    function validId(value) {
        if (!/^\d+$/.test(text(value))) throw Error('Invalid record or line ID.');
        return text(value);
    }

    function findLine(shipment, itemId, inventoryId) {
    function findLine(shipment, itemId, inventoryId, poId) {
        if (!lineMaps.has(shipment)) {
            const count = shipment.getLineCount({sublistId: 'items'});
            const shipmentLines = [];
            for (let line = 0; line < count; line++) {
                shipmentLines.push({line,
                    po: text(shipment.getSublistValue({sublistId: 'items', fieldId: 'purchaseorder', line})),
                    key: text(shipment.getSublistValue({sublistId: 'items', fieldId: 'shipmentitem', line}))});
            }
            const poIds = [...new Set(shipmentLines.map(l => l.po).filter(Boolean))];
            const poItems = {};
            if (poIds.length) {
                const lookup = search.create({type: 'purchaseorder', filters: [['internalid','anyof',poIds], 'AND', ['mainline','is','F']],
                    columns: [search.createColumn({name:'internalid',sort:search.Sort.ASC}), search.createColumn({name:'lineuniquekey',sort:search.Sort.ASC}), 'item']}).run();
                for (let start = 0; ; start += 1000) {
                    const rows = lookup.getRange({start, end: start + 1000});
                    rows.forEach(r => { poItems[text(r.getValue('internalid')) + ':' + text(r.getValue('lineuniquekey'))] = text(r.getValue('item')); });
                    if (rows.length < 1000) break;
                }
            }
            const map = {};
            shipmentLines.forEach(l => {
                let item = poItems[l.po + ':' + l.key];
                if (!item && shipment.hasSublistSubrecord({sublistId:'items',fieldId:'inventorydetail',line:l.line})) {
                    item = text(shipment.getSublistSubrecord({sublistId:'items',fieldId:'inventorydetail',line:l.line}).getValue({fieldId:'item'}));
                }
                if (!item) throw Error('Cannot identify an IBS item line. No inventory details were saved.');
                if (!map[item]) map[item] = [];
                map[item].push(l.line);
            });
            lineMaps.set(shipment, map);
        }
        let matches = lineMaps.get(shipment)[text(itemId)] || [];
        if (poId) matches = matches.filter(line => text(shipment.getSublistValue({sublistId:'items',fieldId:'purchaseorder',line})) === text(poId));
        if (inventoryId) matches = matches.filter(line => text(shipment.getSublistValue({sublistId:'items',fieldId:'inventorydetail',line})) === text(inventoryId));
        if (matches.length > 1) throw Error('This item appears on multiple lines in the shipment. Item alone cannot identify which line to update.');
        if (matches.length > 1) throw Error('This item appears on multiple lines in the shipment. PO and item still match multiple lines. A unique inventory detail is required.');
        if (!matches.length) throw Error('The item is no longer on this shipment. Refresh the page.');
        return matches[0];
    }

    function itemScope(shipmentId, itemId) {
    function itemScope(shipmentId, itemId, poId) {
        const searchId = runtime.getCurrentScript().getParameter({name: PARAM});
        if (!searchId) throw Error('Set the deployment parameter ' + PARAM + '.');
        const saved = search.load({id: searchId});
        const inventoryColumns = saved.columns.filter(c => join(c) === 'inventorydetail').map(c => ({name:c.name,label:c.label || c.name}));
        const itemColumn = search.createColumn({name:'item'});
        saved.columns = [itemColumn];
        saved.filters = saved.filters.concat([
            search.createFilter({name:'internalid',operator:search.Operator.ANYOF,values:shipmentId}),
            search.createFilter({name:'item',operator:search.Operator.ANYOF,values:itemId})]);
        if (poId) saved.filters = saved.filters.concat(search.createFilter({name:'purchaseorder',operator:search.Operator.ANYOF,values:poId}));
        const rows = saved.run().getRange({start:0,end:1});
        if (!rows.length) throw Error('This shipment item is no longer included in the configured search.');
        return {item:text(rows[0].getText(itemColumn) || rows[0].getValue(itemColumn)), columns:inventoryColumns};
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
            rows.push({number: text(label('receiptinventorynumber') || label('issueinventorynumber') || value('receiptinventorynumber')),
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

    function getDetail(params, loadedShipment) {
        const shipmentId = validId(params.shipmentId);
        const lineId = validId(params.lineId);
        const scope = itemScope(shipmentId, lineId);
        const scope = itemScope(shipmentId, lineId, params.poId);
        const shipment = loadedShipment || record.load({type: 'inboundshipment', id: shipmentId});
        const index = findLine(shipment, lineId, params.inventoryId);
        const index = findLine(shipment, lineId, params.inventoryId, params.poId);
        const state = lineState(shipment, index);
        const rules = itemRules(lineId);
        const units = unitInfo(lineId, state.unit);
        const expected = state.expected * units.rate, received = state.received * units.rate;
        const bins = rules.bins && state.locationId ? options('bin', [['location','anyof',state.locationId], 'AND', ['inactive','is','F']], 'binnumber') : [];
        const statuses = rules.statuses ? options('inventorystatus', [['isinactive','is','F']], 'name') : [];
        return {shipmentId, lineId, ...state, expected, received, units, rules, bins, statuses, snapshot: snapshot(state),
            location: text(shipment.getSublistText({sublistId: 'items', fieldId: 'receivinglocation', line: index})),
            item: scope.item,
            max: Math.max(0, expected - received), columns: scope.columns};
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

    function currentSearchSnapshot(shipmentId, poId, itemId, inventoryId) {
        const saved = search.load({id:runtime.getCurrentScript().getParameter({name:PARAM})});
        saved.columns = saved.columns.filter(c => !(join(c) === 'inboundshipmentitem' && ['internalid','id'].includes(c.name)));
        const location = saved.columns.find(c => !join(c) && c.name === 'receivinglocation') || search.createColumn({name:'receivinglocation'});
        if (!saved.columns.includes(location)) saved.columns = saved.columns.concat(location);
        saved.filters = saved.filters.concat([
            search.createFilter({name:'internalid',operator:search.Operator.ANYOF,values:shipmentId}),
            search.createFilter({name:'purchaseorder',operator:search.Operator.ANYOF,values:poId}),
            search.createFilter({name:'item',operator:search.Operator.ANYOF,values:itemId})]);
        const results = saved.run();
        let state;
        for (let start = 0; ; start += 1000) {
            const batch = results.getRange({start,end:start+1000});
            batch.forEach(result => {
                const get = (name, joined) => {
                    const column = saved.columns.find(c => c.name === name && join(c) === (joined || ''));
                    return column ? result.getValue(column) : '';
                };
                if (text(get('internalid','inventorydetail')) !== inventoryId) return;
                if (!state) state = {expected:number(get('quantityexpected')),received:number(get('quantityreceived')),
                    locationId:text(result.getValue(location)),inventory:{rows:[]}};
                const qty = get('quantity','inventorydetail');
                if (qty === '' || qty == null || Number(qty) === 0) return;
                const lotColumn = saved.columns.find(c => c.name === 'inventorynumber' && join(c) === 'inventorydetail');
                const expiry = get('expirationdate','inventorydetail');
                const row = {number:text(lotColumn ? result.getText(lotColumn) || result.getValue(lotColumn) : ''),
                    bin:text(get('binnumber','inventorydetail')),status:text(get('status','inventorydetail')),
                    expiry:expiry ? isoDate(format.parse({value:text(expiry),type:format.Type.DATE})) : '',quantity:Number(qty)};
                if (!state.inventory.rows.some(r => JSON.stringify(r) === JSON.stringify(row))) state.inventory.rows.push(row);
            });
            if (batch.length < 1000) break;
        }
        if (!state) throw Error('The shipment item or inventory detail changed. Refresh the page before submitting.');
        return snapshot(state);
    }

    function writeAssignments(shipment, index, rows, detail) {
        if (!rows.length) {
            if (shipment.hasSublistSubrecord({sublistId:'items',fieldId:'inventorydetail',line:index})) {
                shipment.removeSublistSubrecord({sublistId:'items',fieldId:'inventorydetail',line:index});
            }
            return;
        }
        const identity = row => JSON.stringify([text(row.number),text(row.bin),text(row.status),text(row.expiry)]);
        const pending = rows.slice();
        const retained = [];
        const removed = [];
        detail.inventory.rows.forEach((old, line) => {
            const match = pending.findIndex(row => identity(row) === identity(old));
            if (match < 0) removed.push(line);
            else retained.push({old,row:pending.splice(match,1)[0]});
        });
        const subrecord = shipment.getSublistSubrecord({sublistId:'items',fieldId:'inventorydetail',line:index});
        removed.reverse().forEach(line => subrecord.removeLine({sublistId:'inventoryassignment',line}));
        retained.forEach(({old,row}, line) => {
            if (Number(old.quantity) !== Number(row.quantity)) subrecord.setSublistValue({sublistId:'inventoryassignment',fieldId:'quantity',line,value:Number(row.quantity)});
        });
        pending.forEach((row, offset) => {
            const line = retained.length + offset;
            const set = (fieldId,value) => subrecord.setSublistValue({sublistId:'inventoryassignment',fieldId,line,value});
            if (detail.rules.lot || detail.rules.serial) set('receiptinventorynumber',row.number);
            if (row.expiry) set('expirationdate',parseDate(row.expiry));
            if (detail.rules.bins) set('binnumber',Number(row.bin));
            if (detail.rules.statuses) set('inventorystatus',Number(row.status));
            set('quantity',Number(row.quantity));
        });
    }

    function saveDetails(payload) {
        const shipmentId = validId(payload.shipmentId);
        if (!Array.isArray(payload.lines) || !payload.lines.length) throw Error('No changed item lines to submit.');
        const shipment = record.load({type: 'inboundshipment', id: shipmentId, isDynamic: false});
        const changed = new Set();
        const seen = new Set();
        for (const change of payload.lines) {
            if (runtime.getCurrentScript().getRemainingUsage() < 180) throw Error('Too many edited lines in one submission. Submit fewer lines at a time. No changes were saved for this shipment.');
            const lineId = validId(change.itemId);
            const inventoryId = text(change.inventoryId);
            const poId = validId(change.poId);
            const index = findLine(shipment, lineId, inventoryId);
            const index = findLine(shipment, lineId, inventoryId, poId);
            if (seen.has(index)) throw Error('The same IBS line was edited through multiple search rows. Submit changes from one row for that IBS line.');
            seen.add(index);
            if (change.qcStatus !== undefined) {
                itemScope(shipmentId, lineId);
                itemScope(shipmentId, lineId, poId);
                const desired = text(change.qcStatus);
                if (!['','1','2','3','4','5'].includes(desired)) throw Error('Invalid QC Status.');
                const current = text(shipment.getSublistValue({sublistId:'items',fieldId:QC_FIELD,line:index}));
                if (current !== desired) {
                    if (current !== text(change.qcOriginal)) throw Error('QC Status changed in NetSuite. Refresh the page before submitting.');
                    shipment.setSublistValue({sublistId:'items',fieldId:QC_FIELD,line:index,value:desired});
                    changed.add(index);
                    log.debug({title:'IBS QC Status changed',details:{shipmentId,itemId:lineId,previous:current,status:desired}});
                }
            }
            if (change.rows === undefined) continue;
            const detail = getDetail({shipmentId, lineId, inventoryId}, shipment);
            if (detail.snapshot !== change.snapshot) {
            const detail = getDetail({shipmentId, lineId, inventoryId, poId}, shipment);
            if (currentSearchSnapshot(shipmentId, poId, lineId, inventoryId) !== change.snapshot) {
                throw Error('Shipment quantities or inventory details changed for ' + detail.item + '. Refresh the page before submitting.');
            }
            const total = validateRows(change.rows, detail);
            changed.add(index);
            const subrecord = shipment.getSublistSubrecord({sublistId: 'items', fieldId: 'inventorydetail', line: index});
            for (let i = subrecord.getLineCount({sublistId: 'inventoryassignment'}) - 1; i >= 0; i--) {
                subrecord.removeLine({sublistId: 'inventoryassignment', line: i});
            try {
                writeAssignments(shipment, index, change.rows, detail);
            } catch (error) {
                log.error({title:'IBS assignment update failed',details:{shipmentId,poId,itemId:lineId,line:index,before:detail.inventory.rows.length,after:change.rows.length,error:error.message}});
                throw error;
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
            log.debug({title: 'IBS line validated', details: {shipmentId, lineId, before:detail.inventory.rows.length, rows: change.rows.length, total, maximum: detail.max}});
        }
        if (!changed.size) return {id:shipmentId,lines:0};
        const id = shipment.save({enableSourcing: true, ignoreMandatoryFields: false});
        log.audit({title: 'IBS line changes saved', details: {shipmentId: id, lines: Array.from(changed), userId: runtime.getCurrentUser().id}});
        return {id, lines: changed.size};
    }

    function buildPage() {
        const endpoint = url.resolveScript({scriptId: runtime.getCurrentScript().id, deploymentId: runtime.getCurrentScript().deploymentId});
        return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + TITLE + '</title><style>' + styles() + '</style></head><body>' +
            '<main><header><div><h1>' + TITLE + '</h1><p>Review shipments and validate inventory details</p></div><div><button id="refresh">Refresh</button> <button id="submit" class="primary" disabled>Submit</button></div></header>' +
            '<section class="metrics"><div><b id="shipCount">0</b>Shipments</div><div><b id="lineCount">0</b>Item Lines</div><div><b id="editCount">0</b>Changed Lines</div></section>' +
            '<form id="filters">' + [['ibs','Shipment Number'],['container','Container Number'],['seal','Seal Number']].map(([name,label]) =>
                '<div class="filter-field"><label for="filter-' + name + '">' + label + '</label><input id="filter-' + name + '" name="' + name + '" autocomplete="off" role="combobox" aria-expanded="false" aria-controls="choices-' + name + '" placeholder="Type to search"><div id="choices-' + name + '" class="filter-choices" hidden></div></div>').join('') + '<button type="button" id="clear">Clear</button></form>' +
            '<div id="message" role="status"></div><div class="table-wrap"><table id="shipments"></table></div><footer>Open + to view item lines. Click Inventory Detail to view or edit assignments.</footer></main>' +
            '<div id="modal" class="modal" role="dialog" aria-modal="true" aria-label="Inventory Detail" hidden><div class="dialog"><div class="dialog-head"><strong>Inventory Detail</strong><button id="close">×</button></div><div id="detailBody" class="dialog-body"></div><div class="dialog-actions"><button id="cancel">Cancel</button><button id="ok" class="primary">OK</button></div></div></div>' +
            '<div id="imageModal" class="modal" role="dialog" aria-modal="true" aria-label="Image preview" hidden><div class="image-dialog"><button id="closeImage">Close</button><img id="largeImage" alt="Full size image"></div></div>' +
            '<script>(' + client.toString() + ')(' + JSON.stringify(endpoint).replace(/</g, '\\u003c') + ');</script></body></html>';
    }

    function styles() {
        return `.filter-field{position:relative}.filter-choices{position:absolute;top:100%;left:0;right:0;max-height:250px;overflow:auto;background:white;border:1px solid #bdc7d8;border-radius:5px;box-shadow:0 8px 20px #122d4a22;z-index:5}.filter-choices button{display:block;width:100%;text-align:left;border:0;border-radius:0;font-weight:normal}.filter-choices button:hover{background:#e8eef7}.inv-entry{display:flex;gap:8px;flex-wrap:wrap;align-items:end;margin-bottom:16px}.inv-entry label{flex:1;min-width:120px}.inventory-icon{width:20px;height:20px;object-fit:contain}.detail-button{border:0!important;background:transparent!important;padding:3px!important}*{box-sizing:border-box}body{margin:0;background:#f3f6fa;color:#24364b;font:13px Arial,sans-serif}main{margin:20px;border:1px solid #d7dde7;border-radius:8px;overflow:hidden;background:white}header{background:#122d4a;color:white;padding:22px;display:flex;justify-content:space-between;align-items:center;gap:20px}h1{font-size:22px;margin:0}header p{margin:7px 0 0;color:#c4d2e1}button{cursor:pointer;background:white;color:#24364b;border:1px solid #bdc7d8;border-radius:5px;padding:8px 13px;font-weight:bold}button.primary{background:#1664c0;border-color:#1664c0;color:white}button:disabled{opacity:.5;cursor:default}.metrics{display:flex;gap:14px;background:#f8fafc;padding:18px}.metrics>div{background:white;border:1px solid #dce3ed;border-radius:7px;padding:14px 22px;min-width:160px;color:#64748b}.metrics b{display:block;font-size:25px;color:#163b63;margin-bottom:5px}form{display:flex;gap:12px;padding:16px;align-items:end;border-bottom:1px solid #d7dde7;flex-wrap:wrap}label{display:block;font-size:11px;font-weight:bold}input,select{display:block;margin-top:5px;width:100%;border:1px solid #bdc7d8;border-radius:5px;padding:7px;background:white;color:#24364b}form input{width:230px}#message{padding:12px 16px;white-space:pre-wrap}#message:empty{display:none}.error{color:#a52727;background:#fff0ef}.success{color:#17603c;background:#edf9f2}.table-wrap{overflow:auto;max-height:65vh}table{border-collapse:separate;border-spacing:0;width:100%;font-size:12px}th{background:#e8eef7;color:#24364b;border-bottom:1px solid #cad4e3;border-right:1px solid #d8e0ea;padding:10px;text-align:left;white-space:nowrap}#shipments>thead th{position:sticky;top:0;z-index:1}td{border-bottom:1px solid #e6ebf2;border-right:1px solid #edf1f6;padding:9px;vertical-align:middle;min-width:90px}tr:hover>td{background:#f8fbff}td:first-child{min-width:40px}a{color:#165ba7;text-decoration:none;font-weight:bold}a:hover{text-decoration:underline}.expanded>td{padding:14px;background:#f9fbfd}.item-scroll{overflow:auto;max-width:calc(100vw - 100px)}.item-scroll table{min-width:1120px}.expander{padding:2px;width:24px;height:24px}.thumb{width:45px;height:45px;object-fit:contain;cursor:zoom-in}.detail-button{color:#1664c0;font-size:19px;padding:3px 8px}.filled{color:#168052}.dirty{background:#fff4d5!important}footer{padding:13px 16px;color:#667085}.modal{position:fixed;inset:0;z-index:10;display:flex;align-items:center;justify-content:center;background:rgba(15,23,42,.38)}[hidden]{display:none!important}.dialog{width:min(1100px,calc(100vw - 32px));max-height:calc(100vh - 44px);overflow:auto;background:white;border:1px solid #c9d4e4;border-radius:8px;box-shadow:0 24px 70px rgba(15,23,42,.24)}.dialog-head{display:flex;justify-content:space-between;align-items:center;padding:12px 14px;background:#e8eef7;border-bottom:1px solid #cad4e3}.dialog-head strong{font-size:15px}.dialog-body{padding:16px}.dialog-actions{display:flex;justify-content:flex-end;gap:8px;padding:12px 14px;border-top:1px solid #d7dde7}.detail-summary{display:flex;gap:25px;flex-wrap:wrap;margin-bottom:15px}.detail-summary b{display:block;margin-top:4px}.inventory-wrap{overflow:auto}.inventory-wrap input,.inventory-wrap select{min-width:110px}.inventory-wrap input[type=number]{width:95px;min-width:95px}.image-dialog{background:white;padding:15px;border-radius:8px;max-width:94vw}.image-dialog img{display:block;max-width:90vw;max-height:80vh;margin-top:10px}.hint{color:#667085;padding:10px 0}.empty{text-align:center;padding:40px;color:#64748b}@media(max-width:700px){main{margin:8px}header{align-items:flex-start;flex-direction:column}.metrics{gap:6px}.metrics>div{min-width:0;padding:12px;flex:1}h1{font-size:19px}}`;
    }

    function client(endpoint) {
        const $ = id => document.getElementById(id);
        const escape = value => String(value == null ? '' : value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
        const qcOptions = [{id:'1',name:'Release'},{id:'2',name:'To Be Labelled'},{id:'3',name:'Pending QC Release'},{id:'4',name:'QC Released'},{id:'5',name:'QC DEVIATE'}];
        const selected = {ibs:'',container:'',seal:''};
        const filterFields = {ibs:'shipmentnumber',container:'custrecord157',seal:'custrecord158'};
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
                if (c.isImage && !c.image) return '<span class="hint">Image unavailable</span>';
                if (c.image) return '<img class="thumb" src="' + escape(c.image) + '" data-image="' + escape(c.image) + '" alt="' + escape(c.text) + '">';
                return c.url ? '<a target="_blank" rel="noopener" href="' + escape(c.url) + '">' + escape(c.text) + '</a>' : escape(c.text);
            }).join('<br>');
        }
        function filterValues(shipment, name) {
            const column = data.columns.find(c => !c.join && c.name === filterFields[name]);
            return column ? (shipment.cells[column.key] || []).map(c => c.text) : [];
        }
        function filteredShipments() {
            return data.shipments.filter(s => Object.keys(selected).every(name => !selected[name] || filterValues(s,name).includes(selected[name])));
        }
        function hideChoices() {
            Object.keys(selected).forEach(name => { $('choices-'+name).hidden = true; $('filter-'+name).setAttribute('aria-expanded','false'); });
        }
        function showChoices(name, query) {
            hideChoices();
            const list = $('choices-'+name);
            const values = [...new Set(data.shipments.flatMap(s => filterValues(s,name)).filter(Boolean))]
                .filter(value => value.toLowerCase().includes(query.toLowerCase())).sort((a,b) => a.localeCompare(b,undefined,{numeric:true}));
            list.innerHTML = '';
            [''].concat(values).forEach(value => {
                const button = document.createElement('button'); button.type = 'button'; button.textContent = value || 'All';
                button.onmousedown = event => event.preventDefault();
                button.onclick = () => { selected[name] = value; $('filter-'+name).value = value; hideChoices(); render(); };
                list.appendChild(button);
            });
            list.hidden = false; $('filter-'+name).setAttribute('aria-expanded','true');
        }
        function defaultQcDrafts() {
            data.shipments.forEach(s => s.lines.forEach(l => {
                if (l.qcEnabled && l.qcInitial !== l.qcOriginal) {
                    const k = key(s.id,l.id);
                    if (!edits[k]) edits[k] = {shipmentId:s.id,lineId:l.id,itemId:l.itemId,inventoryId:l.detail.inventory.id,qcStatus:l.qcInitial,qcOriginal:l.qcOriginal};
                    if (!edits[k]) edits[k] = {shipmentId:s.id,lineId:l.id,itemId:l.itemId,poId:l.poId,inventoryId:l.detail.inventory.id,qcStatus:l.qcInitial,qcOriginal:l.qcOriginal};
                }
            }));
        }
        function mergeChanges(lines) {
            const merged = {};
            lines.forEach(line => {
                const k = line.itemId + ':' + line.inventoryId;
                const k = line.poId + ':' + line.itemId + ':' + line.inventoryId;
                if (!merged[k]) { merged[k] = {...line}; return; }
                const target = merged[k];
                if (line.qcStatus !== undefined) {
                    if (target.qcStatus !== undefined && target.qcStatus !== line.qcStatus) throw Error('Conflicting QC Status values on rows for the same IBS item line.');
                    target.qcStatus = line.qcStatus; target.qcOriginal = line.qcOriginal;
                }
                if (line.rows !== undefined) {
                    if (target.rows !== undefined && JSON.stringify(target.rows) !== JSON.stringify(line.rows)) throw Error('Conflicting inventory edits on rows for the same IBS item line.');
                    target.rows = line.rows; target.snapshot = line.snapshot;
                }
            });
            return Object.values(merged);
        }
        function render() {
            const shown = filteredShipments();
            const headers = data.columns.filter(c => c.section === 'header');
            const items = data.columns.filter(c => c.section === 'item');
            const remaining = items.findIndex(c => c.name === 'quantityremaining');
            items.splice(remaining >= 0 ? remaining + 1 : items.length,0,{key:'inventoryButton',label:'Inventory Detail'});
            let html = '<thead><tr><th></th>' + headers.map(c => '<th>' + escape(c.label) + '</th>').join('') + '</tr></thead><tbody>';
            shown.forEach(s => {
                html += '<tr><td><button class="expander" aria-expanded="' + expanded.has(s.id) + '" data-expand="' + s.id + '">' + (expanded.has(s.id) ? '−' : '+') + '</button></td>' + headers.map(c => '<td>' + cell(s.cells[c.key]) + '</td>').join('') + '</tr>';
                if (expanded.has(s.id)) {
                    html += '<tr class="expanded"><td colspan="' + (headers.length + 1) + '"><div class="item-scroll"><table><thead><tr>' + items.map(c => '<th>' + escape(c.label) + '</th>').join('') + '</tr></thead><tbody>';
                    s.lines.forEach(l => {
                        const draft = edits[key(s.id,l.id)];
                        html += '<tr>' + items.map(c => {
                            if (c.key === 'inventoryButton') return '<td class="' + (draft ? 'dirty' : '') + '"><button class="detail-button ' + ((draft && draft.rows ? draft.rows.length : l.hasDetail) ? 'filled' : '') + '" title="View / Edit Inventory Detail" aria-label="View / Edit Inventory Detail" data-detail="' + s.id + ':' + l.id + '">' + '<img class="inventory-icon" alt="Inventory Detail" src="' + ((draft && draft.rows ? draft.rows.length : l.hasDetail) ? 'https://4382108.app.netsuite.com/core/media/media.nl?id=24230&c=4382108&h=IH_6SQ4VYeAu0pFkOMmf5qXj8CSBZWAU0A5XLbIcoYkJkseL' : 'https://4382108.app.netsuite.com/core/media/media.nl?id=24231&c=4382108&h=YnSYg6zHZBKBjFQ6yI7HCuuSDbzV1x356tga3ZAREJ8Ix3f3') + '">' + '</button></td>';
                            if (c.name === 'custrecord_mi_qc_status') {
                                const value = draft && draft.qcStatus !== undefined ? draft.qcStatus : l.qcInitial;
                                return '<td><select aria-label="QC Status" data-qc="' + s.id + ':' + l.id + '"' + (busy ? ' disabled' : '') + '>' + options(qcOptions,value) + '</select></td>';
                            }
                            return '<td>' + cell(l.cells[c.key]) + '</td>';
                        }).join('') + '</tr>';
                    });
                    html += '</tbody></table></div></td></tr>';
                }
            });
            html += '</tbody>';
            if (!shown.length) html += '<tbody><tr><td class="empty" colspan="' + (headers.length+1) + '">No matching shipments.</td></tr></tbody>';
            $('shipments').innerHTML = html;
            $('shipCount').textContent = shown.length;
            $('lineCount').textContent = shown.reduce((sum,s) => sum+s.lines.length,0);
            $('editCount').textContent = Object.keys(edits).length;
            setBusy(busy);
        }
        async function load(discard) {
            if (busy) return;
            if (discard && Object.keys(edits).length && !confirm('Discard pending line changes and refresh?')) return;
            if (discard) edits = {};
            setBusy(true); message('Loading shipments…');
            try { data = await request('list'); defaultQcDrafts(); message(''); render(); }
            catch(error) { message(error.message,true); }
            finally { setBusy(false); }
        }
        function options(list, value) {
            let html = '<option value="">Select</option>';
            if (value && !list.some(v => v.id === value)) html += '<option selected value="' + escape(value) + '">Unavailable (' + escape(value) + ')</option>';
            return html + list.map(v => '<option value="' + escape(v.id) + '"' + (v.id === value ? ' selected' : '') + '>' + escape(v.name) + '</option>').join('');
        }
        function inventoryCell(column, row) {
            switch(column.name) {
                case 'internalid': return escape(active.inventory.id);
                case 'item': return escape(active.item);
                case 'location': return escape(active.location);
                case 'inventorynumber': return escape(row.number);
                case 'binnumber': return escape((active.bins.find(b => b.id === row.bin) || {}).name || row.bin);
                case 'status': return escape((active.statuses.find(v => v.id === row.status) || {}).name || row.status);
                case 'expirationdate': return escape(row.expiry);
                case 'quantity': return escape(row.quantity);
                default: return '—';
            }
        }
        function showDetail() {
            const required = [['inventorynumber','Number'],['binnumber','Bin Number'],['expirationdate','Expiration Date'],['quantity','Quantity'],['status','Status']];
            const columns = active.columns.slice();
            required.forEach(([name,label]) => { if (!columns.some(c => c.name === name)) columns.push({name,label}); });
            const label = (name,fallback) => escape((columns.find(c => c.name === name) || {}).label || fallback);
            const numberField = active.rules.lot || active.rules.serial ? '<label>' + label('inventorynumber','Serial/Lot Number') + '<input id="inv-number"></label>' : '';
            const expiryField = active.rules.lot ? '<label>' + label('expirationdate','Expiry Date') + '<input id="inv-expiry" type="date"></label>' : '';
            const binField = active.rules.bins ? '<label>' + label('binnumber','Bin') + '<select id="inv-bin">' + options(active.bins,'') + '</select></label>' : '';
            const statusField = active.rules.statuses ? '<label>' + label('status','Status') + '<select id="inv-status">' + options(active.statuses,'') + '</select></label>' : '';
            $('detailBody').innerHTML = '<div class="hint">Quantities in ' + escape(active.units.name || 'base units') + '</div><div class="detail-summary"><div>Item<b>' + escape(active.item) + '</b></div><div>Qty Expected<b>' + active.expected + '</b></div><div>Qty Received<b>' + active.received + '</b></div><div>Maximum Quantity<b>' + active.max + '</b></div><div>Total Qty<b id="total"></b></div></div>' +
                '<div class="inv-entry">' + numberField + expiryField + binField + statusField + '<label>' + label('quantity','Quantity') + '<input id="inv-qty" type="number" min="0" step="any"></label><button class="primary" id="addRow">Add Row</button></div>' +
                '<div class="inventory-wrap"><table><thead><tr>' + columns.map(c => '<th>' + escape(c.label) + '</th>').join('') + '<th></th></tr></thead><tbody>' + active.rows.map((row,i) => '<tr>' + columns.map(c => '<td>' + inventoryCell(c,row) + '</td>').join('') + '<td><button data-edit="' + i + '">Edit</button> <button data-remove="' + i + '">Remove</button></td></tr>').join('') + '</tbody></table></div><div class="hint">Total Qty must not exceed Qty Expected − Qty Received. OK stages changes; Submit saves the shipment.</div><div id="detailError" class="error" role="alert"></div>';
            active.editIndex = null;
            updateTotal();
        }
        function addInventoryRow() {
            const value = id => $(id) ? $(id).value.trim() : '';
            const row = {number:value('inv-number'),expiry:value('inv-expiry'),bin:value('inv-bin'),status:value('inv-status'),quantity:Number(value('inv-qty'))};
            let error = '';
            if (!Number.isFinite(row.quantity) || row.quantity <= 0) error = 'Enter Quantity greater than zero.';
            if ((active.rules.lot || active.rules.serial) && !row.number) error = 'Enter Serial/Lot Number.';
            if (active.rules.bins && !row.bin) error = 'Select Bin.';
            if (active.rules.statuses && !row.status) error = 'Select Status.';
            if (active.rules.serial && row.quantity !== 1) error = 'Serial quantity must be 1.';
            if ($('inv-expiry') && !$('inv-expiry').checkValidity()) error = 'Enter a valid expiration date.';
            const otherRows = active.rows.filter((r,i) => i !== active.editIndex);
            const total = otherRows.reduce((sum,r) => sum+Number(r.quantity),0) + row.quantity;
            if (total-active.max > 0.00000001) error = 'Inventory detail quantity cannot be more than ' + active.max + '.';
            if (active.rules.serial && otherRows.some(r => r.number.toLowerCase() === row.number.toLowerCase())) error = 'This serial number already exists.';
            if (error) { $('detailError').textContent = error; return; }
            if (error) { $('detailError').textContent = error; return false; }
            const duplicate = active.rules.lot && active.editIndex === null ? active.rows.find(r => r.number.toLowerCase() === row.number.toLowerCase() && r.bin === row.bin && r.status === row.status && r.expiry === row.expiry) : null;
            if (duplicate) {
                if (!confirm('This lot number already exists. Click OK to merge and add the quantity.')) return;
                duplicate.quantity = Number(duplicate.quantity) + row.quantity;
            } else if (active.editIndex !== null) active.rows[active.editIndex] = row;
            else active.rows.push(row);
            showDetail();
            return true;
        }
        function updateTotal() { $('total').textContent = Math.round(active.rows.reduce((sum,r) => sum + (Number(r.quantity)||0),0)*1e8)/1e8; }
        function openDetail(shipmentId,lineId) {
            if (busy) return;
            const shipment = data.shipments.find(s => s.id === shipmentId);
            const line = shipment && shipment.lines.find(l => l.id === lineId);
            if (!line || !line.detail) { message('Inventory detail is not available. Refresh the page.',true); return; }
            const stored = line.detail;
            const draft = edits[key(shipmentId,lineId)];
            active = {...stored,rows:JSON.parse(JSON.stringify(draft && draft.rows ? draft.rows : stored.inventory.rows))};
            showDetail(); $('modal').hidden = false; $('close').focus();
        }
        function closeDetail() { $('modal').hidden = true; active = null; }
        function stageDetail() {
            if (active.editIndex !== null && !addInventoryRow()) return;
            if (['inv-number','inv-expiry','inv-bin','inv-status','inv-qty'].some(id => $(id) && $(id).value)) { $('detailError').textContent = 'Click Add Row / Update Row before OK, or clear the entry fields.'; return; }
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
                if (error) { $('detailError').textContent = error; return false; }
            }
            if (total-active.max > 0.00000001) { $('detailError').textContent = 'Total inventory quantity cannot exceed ' + active.max + '.'; return; }
            const k = key(active.shipmentId,active.lineId);
            const draft = edits[k] || {shipmentId:active.shipmentId,lineId:active.lineId,itemId:active.itemId,inventoryId:active.inventory.id};
            const draft = edits[k] || {shipmentId:active.shipmentId,lineId:active.lineId,itemId:active.itemId,poId:active.poId,inventoryId:active.inventory.id};
            if (JSON.stringify(active.rows) === JSON.stringify(active.inventory.rows)) { delete draft.rows; delete draft.snapshot; }
            else { draft.rows = active.rows; draft.snapshot = active.snapshot; }
            if (draft.rows !== undefined || draft.qcStatus !== undefined) edits[k] = draft;
            else delete edits[k];
            closeDetail(); render();
        }
        async function submit() {
            if (busy || !Object.keys(edits).length) return;
            setBusy(true); message('Validating and saving changed lines…');
            const groups = {};
            Object.values(edits).forEach(e => (groups[e.shipmentId] ||= []).push(e));
            let saved = 0;
            try {
                for (const [shipmentId,lines] of Object.entries(groups)) {
                    await request('save', {}, {shipmentId,lines:mergeChanges(lines)});
                    lines.forEach(l => delete edits[key(shipmentId,l.lineId)]);
                    saved++;
                }
                data = await request('list');
                defaultQcDrafts();
                message('Line changes saved for ' + saved + ' shipment(s).');
            } catch(error) { message((saved ? saved + ' shipment(s) saved. ' : '') + error.message + ' Remaining drafts are retained. Refresh the page if the record changed.',true); }
            finally { setBusy(false); render(); }
        }
        $('filters').onsubmit = event => event.preventDefault();
        Object.keys(selected).forEach(name => {
            const input = $('filter-'+name);
            input.onfocus = () => showChoices(name,'');
            input.oninput = () => { if (!input.value) { selected[name] = ''; render(); } showChoices(name,input.value); };
            input.onkeydown = event => {
                if (event.key === 'Escape') { hideChoices(); input.value = selected[name]; }
                if (event.key === 'Enter') { event.preventDefault(); const choices = $('choices-'+name).querySelectorAll('button'); if (choices.length === 2) choices[1].click(); }
            };
            input.onblur = () => { input.value = selected[name]; };
        });
        document.addEventListener('click', event => { if (!event.target.closest('.filter-field')) hideChoices(); });
        $('refresh').onclick = () => load(true);
        $('clear').onclick = () => { Object.keys(selected).forEach(name => { selected[name] = ''; $('filter-'+name).value = ''; }); hideChoices(); render(); };
        $('submit').onclick = submit;
        $('shipments').onchange = event => {
            const target = event.target.dataset.qc;
            if (!target || busy) return;
            const [shipmentId,lineId] = target.split(':');
            const shipment = data.shipments.find(s => s.id === shipmentId);
            const line = shipment.lines.find(l => l.id === lineId);
            shipment.lines.filter(l => l.itemId === line.itemId && l.detail.inventory.id === line.detail.inventory.id).forEach(l => {
            shipment.lines.filter(l => l.itemId === line.itemId && l.poId === line.poId && l.detail.inventory.id === line.detail.inventory.id).forEach(l => {
                const k = key(shipmentId,l.id);
                const draft = edits[k] || {shipmentId,lineId:l.id,itemId:l.itemId,inventoryId:l.detail.inventory.id};
                const draft = edits[k] || {shipmentId,lineId:l.id,itemId:l.itemId,poId:l.poId,inventoryId:l.detail.inventory.id};
                l.qcInitial = event.target.value;
                if (event.target.value === l.qcOriginal) { delete draft.qcStatus; delete draft.qcOriginal; }
                else { draft.qcStatus = event.target.value; draft.qcOriginal = l.qcOriginal; }
                if (draft.rows !== undefined || draft.qcStatus !== undefined) edits[k] = draft;
                else delete edits[k];
            });
            render();
        };
        $('shipments').onclick = event => {
            const button = event.target.closest('button');
            if (button && !busy) {
                if (button.dataset.expand) { const id = button.dataset.expand; expanded.has(id) ? expanded.delete(id) : expanded.add(id); render(); }
                if (button.dataset.detail) openDetail(...button.dataset.detail.split(':'));
            }
            if (event.target.dataset.image) { $('largeImage').src = event.target.dataset.image; $('imageModal').hidden = false; $('closeImage').focus(); }
        };
        $('detailBody').onclick = event => {
            if (event.target.id === 'addRow') addInventoryRow();
            if (event.target.dataset.remove !== undefined) { active.rows.splice(Number(event.target.dataset.remove),1); showDetail(); }
            if (event.target.dataset.edit !== undefined) {
                active.editIndex = Number(event.target.dataset.edit); const row = active.rows[active.editIndex];
                [['inv-number','number'],['inv-expiry','expiry'],['inv-bin','bin'],['inv-status','status'],['inv-qty','quantity']].forEach(([id,field]) => { if ($(id)) $(id).value = row[field]; });
                $('addRow').textContent = 'Update Row';
            }
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