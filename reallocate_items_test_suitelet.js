/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * Unsupported reallocateitem adapter. Requires sandbox validation before production.
 * Native browser handoff: no record.save() calls. Self-contained Suitelet with browser-controlled native popup. No UE or separate client file.
 */
define(['N/record', 'N/search', 'N/ui/serverWidget', 'N/runtime', 'N/cache', 'N/crypto/random', 'N/log', 'N/url', 'N/redirect'],
(record, search, ui, runtime, cache, random, log, url, redirect) => {
    const HANDOFF_JS = "    const EPS = 0.000001;\n    let targetWindow;\n    let nativeDocument;\n    let running = false;\n    const wait = ms => new Promise(resolve => setTimeout(resolve, ms));\n    const number = v => { const n = Number(v || 0); if (!Number.isFinite(n)) throw Error('Invalid native quantity.'); return n; };\n    const field = (name, line) => targetWindow.nlapiGetLineItemValue('order', name, line);\n    function rows() {\n        const result = [];\n        for (let line = 1; line <= targetWindow.nlapiGetLineItemCount('order'); line++) {\n            const key = [String(field('transactiontype', line) || field('transactiontypename', line) || ''),\n                String(field('orderid', line)), String(field('orderline', line) == null ? '' : field('orderline', line))].join('|');\n            if (result.some(r => r.key === key)) throw Error('Ambiguous native order line ' + key);\n            result.push({ key, line, committed: number(field('quantitycommitted', line)), remaining: number(field('quantityremaining', line)),\n                picked: number(field('quantitypicked', line)), reservation: number(field('quantityreservationcommitted', line)),\n                special: number(field('quantityspecialcommitted', line)), version: String(field('version', line) || '') });\n        }\n        return result;\n    }\n    function banner(message, error) {\n        let box = document.getElementById('rr_handoff_status');\n        if (!box) { box = document.createElement('div'); box.id = 'rr_handoff_status'; document.body.prepend(box); }\n        box.style.cssText = 'padding:16px;margin:12px;border:2px solid ' + (error ? '#a22' : '#267') + ';background:white;font:15px sans-serif;';\n        box.textContent = message;\n    }\n    async function request(endpoint, method = 'GET') {\n        const response = await fetch(endpoint, { method, credentials: 'same-origin', cache: 'no-store' });\n        if (!response.ok) throw Error('Handoff request failed: HTTP ' + response.status);\n        const body = await response.json();\n        if (!body.ok) throw Error(body.error || 'Handoff rejected.');\n        return body;\n    }\n    async function ready() {\n        for (let i = 0; i < 100; i++) {\n            if (targetWindow.NS && targetWindow.NS.form && targetWindow.NS.form.isInited() && targetWindow.NS.form.isValid()) return;\n            await wait(200);\n        }\n        throw Error('Native page did not finish loading. Nothing submitted.');\n    }\n    async function run(data) {\n        const sessionKey = 'rr_native_submit_' + data.token;\n        if (sessionStorage.getItem(sessionKey)) throw Error('This allocation already reached submission in this tab. Inspect commitments before retrying.');\n        targetWindow = window.open('/app/accounting/transactions/reallocitems.nl', '_blank');\n        if (!targetWindow) throw Error('Popup blocked. Allow popups for NetSuite, then click Open Native Page and Apply.');\n        let loaded = false;\n        for (let i = 0; i < 150; i++) {\n            if (targetWindow.closed) throw Error('Native window was closed.');\n            try {\n                nativeDocument = targetWindow.document;\n                if (nativeDocument.readyState === 'complete' && typeof targetWindow.orderSyncRow === 'function') { loaded = true; break; }\n            } catch (_) { /* navigation not ready */ }\n            await wait(200);\n        }\n        if (!loaded) throw Error('Native page did not load or the browser blocked access. Nothing submitted.');\n        const required = ['nlapiSetFieldValue', 'nlapiGetFieldValue', 'nlapiGetLineItemValue', 'nlapiGetLineItemCount',\n            'nlapiSelectLineItem', 'nlapiSetCurrentLineItemValue', 'orderSyncRow', 'orderRecalcMachine'];\n        required.forEach(name => { if (typeof targetWindow[name] !== 'function') throw Error('Native page function unavailable: ' + name); });\n        banner('Loading the approved allocation plan\u2026');\n        const response = await request(data.endpoint);\n        const plan = response.plan;\n        await ready();\n        // Use native field sourcing. Never fetch or replay the native POST body.\n        targetWindow.nlapiSetFieldValue('item', String(plan.config.item), true, true);\n        await ready();\n        targetWindow.nlapiSetFieldValue('location', String(plan.config.location), true, true);\n        await ready();\n        let before, previous = '';\n        for (let i = 0; i < 100; i++) {\n            await wait(200); await ready();\n            before = rows();\n            const signature = JSON.stringify(before);\n            if (signature === previous && before.some(r => r.key === plan.target.key)) break;\n            previous = signature;\n            if (i === 99) throw Error('Destination line did not finish loading. Nothing submitted.');\n        }\n        if (String(targetWindow.nlapiGetFieldValue('item')) !== String(plan.config.item) ||\n            String(targetWindow.nlapiGetFieldValue('location')) !== String(plan.config.location)) throw Error('Native item/location did not match the plan.');\n        const expectedGrid = JSON.parse(plan.snapshot);\n        if (before.length !== expectedGrid.length) throw Error('Native row count changed. Create a new preview.');\n        for (const v of expectedGrid) {\n            const key = [String(v.transactiontype || v.transactiontypename || ''), String(v.orderid), String(v.orderline == null ? '' : v.orderline)].join('|');\n            const r = before.find(x => x.key === key);\n            if (!r || Math.abs(r.committed - number(v.quantitycommitted)) > EPS ||\n                Math.abs(r.remaining - number(v.quantityremaining)) > EPS || Math.abs(r.picked - number(v.quantitypicked)) > EPS ||\n                Math.abs(r.reservation - number(v.quantityreservationcommitted)) > EPS || Math.abs(r.special - number(v.quantityspecialcommitted)) > EPS ||\n                r.version !== String(v.version || '')) throw Error('Native quantities or versions changed since preview. Nothing submitted.');\n        }\n        const changes = [...plan.sources.filter(s => s.taken > EPS), plan.target];\n        banner('Applying the approved quantities to the native page\u2026');\n        for (const change of changes) {\n            const r = before.find(x => x.key === change.key);\n            if (!r) throw Error('Selected row is missing. Nothing submitted.');\n            targetWindow.nlapiSelectLineItem('order', r.line);\n            targetWindow.nlapiSetCurrentLineItemValue('order', 'quantitycommitted', String(change.after), false);\n            // This is the same synchronization path called by native quantity edits.\n            targetWindow.orderSyncRow(r.line, false, false, 'quantitycommitted');\n            targetWindow.orderRecalcMachine();\n        }\n        const after = rows();\n        for (const r of before) {\n            const changed = changes.find(c => c.key === r.key);\n            const expected = changed ? changed.after : r.committed;\n            const actual = after.find(a => a.key === r.key);\n            if (!actual || Math.abs(expected - actual.committed) > EPS) throw Error('Native validation changed the requested quantity for ' + r.key + '. Nothing submitted; reload before proceeding.');\n        }\n        // Avoid native automatic allocation of additional free stock on submit.\n        if (number(targetWindow.nlapiGetFieldValue('quantityuncommitted')) > EPS) throw Error('Native page has uncommitted stock and may auto-allocate extra quantities. Review manually; automatic submission stopped.');\n        const submit = nativeDocument.getElementById('submitter');\n        const form = nativeDocument.forms.main_form;\n        if (!submit || !form || !form.elements.whence) throw Error('Native Submit/return controls are unavailable. Nothing submitted.');\n        // Claim once before clicking. Cache is not an atomic lock; do not run\n        // simultaneous submissions for the same item/location.\n        await request(response.armUrl, 'POST');\n        sessionStorage.setItem(sessionKey, 'submitted');\n        form.elements.whence.value = response.verifyUrl;\n        banner('Submitting through NetSuite. Complete any native validation prompt. If navigation stops, do not submit again without checking commitments.');\n        submit.click(); // Normal handler, validation, serialization, and native form POST.\n        banner('Native Submit was clicked. Check the native window for validation prompts and the final verification result.');\n    }\n\n    document.getElementById('rr_open_native').addEventListener('click', function () {\n        if (running) return;\n        running = true;\n        this.disabled = true;\n        run(HANDOFF_DATA).catch(error => {\n            banner('Allocation stopped: ' + error.message + ' Reload or close the native window before retrying.', true);\n            if (!targetWindow) { running = false; document.getElementById('rr_open_native').disabled = false; }\n            // No automatic resubmission after any failure.\n        });\n    });\n";
    const PREFIX = 'custpage_';
    const EPS = 0.000001;
    const esc = v => String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    const num = v => { const n = Number(v || 0); if (!Number.isFinite(n)) throw Error('Invalid quantity: ' + v); return n; };
    const round = n => Math.round(n * 1000000) / 1000000;
    const yes = v => v === true || v === 'T';
    const store = () => cache.getCache({ name: 'single_suitelet_native_handoff_v1', scope: cache.Scope.PRIVATE });
    const owner = () => String(runtime.getCurrentUser().id) + ':' + String(runtime.getCurrentUser().role);
    function text(rec, fieldId) { try { return rec.getText({ fieldId }) || ''; } catch (_) { return ''; } }
    function val(rec, fieldId) { return rec.getValue({ fieldId }); }
    function input(form, id, label, type, value, source) {
        const f = form.addField({ id: PREFIX + id, label, type, ...(source ? { source } : {}) });
        if (value !== undefined && value !== null) f.defaultValue = String(value);
        return f;
    }
    function hidden(form, id, value) { return input(form, id, id, ui.FieldType.LONGTEXT, value).updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN }); }
    function html(form, id, value) { input(form, id, 'Details', ui.FieldType.INLINEHTML, value); }
    function table(headers, rows) {
        return '<div style="overflow:auto"><table style="border-collapse:collapse" border="1" cellpadding="7"><thead><tr>' +
            headers.map(h => '<th>' + esc(h) + '</th>').join('') + '</tr></thead><tbody>' +
            rows.map(r => '<tr>' + r.map(c => '<td>' + esc(c) + '</td>').join('') + '</tr>').join('') + '</tbody></table></div>';
    }
    function native(item, location) {
        const rec = record.create({ type: 'reallocateitem', isDynamic: true });
        rec.setValue({ fieldId: 'item', value: Number(item) });
        rec.setValue({ fieldId: 'location', value: Number(location) });
        const fields = rec.getSublistFields({ sublistId: 'order' });
        for (const f of ['quantitycommitted', 'quantityremaining', 'commit']) {
            if (!fields.includes(f)) throw Error('Native reallocation does not expose required field ' + f);
        }
        // Only identity fields are used to map rows. Never match arbitrary
        // numeric cells (quantities, dates, priorities) to transaction IDs.
        const override = runtime.getCurrentScript().getParameter({ name: 'custscript_rr_identity_field' });
        const names = ['id', 'internalid', 'orderid', 'doc', 'docid', 'transaction', 'transactionid', 'order', 'ordernumber', 'orderno', 'tranid'];
        const identity = override ? [String(override)] : fields.filter(f => names.includes(f.toLowerCase()));
        if (!identity.length || identity.some(f => !fields.includes(f))) {
            throw Error('Cannot identify native order rows safely. Exposed fields: ' + fields.join(', ') +
                '. Set optional script parameter custscript_rr_identity_field to the verified transaction identity field.');
        }
        const rows = [];
        const count = rec.getLineCount({ sublistId: 'order' });
        if (count > 1000) throw Error('More than 1,000 native rows. Narrow the item/location or extend pagination before use.');
        for (let line = 0; line < count; line++) {
            const values = {}, labels = {};
            fields.forEach(fieldId => {
                values[fieldId] = rec.getSublistValue({ sublistId: 'order', fieldId, line });
                if (identity.includes(fieldId) || fieldId === 'saleschannel') {
                    try { labels[fieldId] = rec.getSublistText({ sublistId: 'order', fieldId, line }); } catch (_) { /* value is sufficient */ }
                }
            });
            rows.push({ line, values, labels });
        }
        return { rec, fields, identity, rows, units: text(rec, 'units') || String(val(rec, 'units') || '') };
    }
    function rowKey(row) {
        const v = row.values;
        if (!/^\d+$/.test(String(v.orderid || ''))) throw Error('Native row has no valid orderid.');
        return [String(v.transactiontype || v.transactiontypename || ''), String(v.orderid), String(v.orderline == null ? '' : v.orderline)].join('|');
    }
    function orderRows(n) {
        const seen = new Set();
        return n.rows.map(row => {
            const key = rowKey(row);
            if (seen.has(key)) throw Error('Native order/line identity is ambiguous: ' + key);
            seen.add(key);
            const v = row.values;
            return { key, id: String(v.orderid), name: String(v.ordernumber || v.orderid),
                type: String(v.transactiontypename || v.transactiontype || ''),
                isReservation: /reserv|ordres/i.test(String(v.transactiontype) + ' ' + String(v.transactiontypename)),
                isSO: /sales.?order|salesord/i.test(String(v.transactiontype) + ' ' + String(v.transactiontypename)),
                channel: String(v.saleschannel || ''), channelText: String(row.labels.saleschannel || v.saleschannel || ''),
                committed: num(v.quantitycommitted), remaining: num(v.quantityremaining), picked: num(v.quantitypicked),
                line: row.line, orderline: String(v.orderline == null ? '' : v.orderline), row };
        });
    }
    function getOrder(n, key) {
        const rows = orderRows(n).filter(r => r.key === key);
        if (rows.length !== 1) throw Error('Selected order line is no longer available: ' + key);
        const r = rows[0], v = r.row.values;
        r.demand = 0;
        if (r.isReservation) {
            const loaded = record.load({ type: 'orderreservation', id: r.id });
            if (String(val(loaded, 'item')) !== String(val(n.rec, 'item')) ||
                String(val(loaded, 'location')) !== String(val(n.rec, 'location'))) throw Error('Reservation item/location mismatch.');
            // Do not mix reservation demand units with native units unless the
            // native and reservation committed quantities agree.
            if (Math.abs(num(val(loaded, 'quantitycommitted')) - r.committed) > EPS) throw Error('Reservation units/commitment differ from the native grid.');
            r.demand = num(val(loaded, 'allocateddemand'));
            if (yes(val(loaded, 'closed'))) throw Error('Reservation ' + r.name + ' is closed.');
        }
        r.available = Math.max(0, r.committed - r.picked - r.demand - num(v.quantityspecialcommitted));
        if (yes(v.isspecialorder) || yes(v.specialorder)) r.available = 0;
        delete r.row;
        return r;
    }
    function snapshot(n) {
        // sys_id and sys_parentid are transient UI row IDs, not transaction IDs.
        return JSON.stringify(n.rows.map(r => {
            const v = { ...r.values }; delete v.sys_id; delete v.sys_parentid;
            return v;
        }));
    }
    function createPlan(c, targetKey, sourceKeys, requested) {
        if (!sourceKeys.length || sourceKeys.length > 20) throw Error('Select between 1 and 20 source order lines.');
        if (new Set(sourceKeys).size !== sourceKeys.length || sourceKeys.includes(targetKey)) throw Error('Sources must be distinct and different from the destination line.');
        const n = native(c.item, c.location);
        const target = getOrder(n, targetKey);
        if (!target.isSO) throw Error('Destination must be a sales order line.');
        if (c.so && target.id !== c.so) throw Error('Destination must be the sales order from which this page was opened.');
        const gap = Math.max(0, target.remaining - target.committed);
        const limit = requested === '' ? gap : num(requested);
        if (limit <= 0) throw Error('Requested quantity must be greater than zero.');
        let needed = Math.min(gap, limit);
        const sources = [];
        // Native page order is the visible consumption order.
        const selected = sourceKeys.map(key => getOrder(n, key)).sort((a, b) => a.line - b.line);
        for (const src of selected) {
            if (src.id === target.id) throw Error('Choose sources from another order, not another line of the destination order.');
            const taken = round(Math.min(needed, src.available));
            needed = round(needed - taken);
            sources.push({ ...src, taken, after: round(src.committed - taken) });
        }
        const moved = round(sources.reduce((sum, src) => sum + src.taken, 0));
        if (moved <= EPS) throw Error('No releasable commitment on selected sources, or destination has no shortfall.');
        return { config: c, target: { ...target, after: round(target.committed + moved) }, sources,
            requested, moved, gap, unmet: round(Math.min(gap, limit) - moved), units: n.units,
            snapshot: snapshot(n), native: n };
    }
    function breakdown(p) {
        return '<p><b>Quantity to move: ' + esc(p.moved) + ' ' + esc(p.units) + '</b>. Unfilled portion of this request: ' + esc(p.unmet) + '.</p>' +
            table(['Source order / line', 'Type', 'Before', 'Taken', 'After', 'Destination SO / line', 'Destination channel'],
                p.sources.map(s => [s.name + ' / ' + s.orderline, s.type, s.committed, s.taken, s.after, p.target.name + ' / ' + p.target.orderline, p.target.channelText || p.target.channel])) +
            '<p>Destination commitment: <b>' + esc(p.target.committed) + ' → ' + esc(p.target.after) + '</b>.</p>';
    }
    function getPlan(key) {
        if (!/^[a-f0-9-]{36}$/i.test(key)) throw Error('Invalid allocation token.');
        const raw = store().get({ key });
        if (!raw) throw Error('Allocation preview expired. Return to the Suitelet.');
        const plan = JSON.parse(raw);
        if (plan.owner !== owner() || Date.now() > plan.expires) throw Error('Allocation preview is expired or belongs to another user/role.');
        return plan;
    }
    function validatePlan(plan) {
        const current = createPlan(plan.config, plan.target.key, plan.sources.map(s => s.key), plan.requested);
        if (current.snapshot !== plan.snapshot || JSON.stringify(current.target) !== JSON.stringify(plan.target) ||
            JSON.stringify(current.sources) !== JSON.stringify(plan.sources)) throw Error('Allocation changed since preview. Create a new preview.');
    }
    function selfUrl(params) {
        const script = runtime.getCurrentScript();
        return url.resolveScript({ scriptId: script.id, deploymentId: script.deploymentId, params });
    }
    function jsonResponse(ctx, data) {
        ctx.response.setHeader({ name: 'Content-Type', value: 'application/json; charset=utf-8' });
        ctx.response.write(JSON.stringify(data));
    }
    function config(p) {
        const c = { item: String(p.custpage_item || ''), location: String(p.custpage_location || ''), so: String(p.soid || p.custpage_so || ''), channel: '' };
        if (c.so) {
            const so = record.load({ type: 'salesorder', id: c.so });
            c.channel = String(val(so, 'saleschannel') || '');
            if (c.item && c.location) {
                let found = false;
                for (let line = 0; line < so.getLineCount({ sublistId: 'item' }); line++) {
                    const item = so.getSublistValue({ sublistId: 'item', fieldId: 'item', line });
                    const loc = so.getSublistValue({ sublistId: 'item', fieldId: 'location', line }) || val(so, 'location');
                    if (String(item) === c.item && String(loc) === c.location) found = true;
                }
                if (!found) throw Error('The selected item/location is not on this SO.');
            }
        }
        return c;
    }
    function onRequest(ctx) {
        const p = ctx.request.parameters;
        if (p.rr_api) {
            try {
                const key = String(p.rr_token || '');
                const plan = getPlan(key);
                if (plan.phase !== 'ready') throw Error('This plan has already reached submission. Do not replay it.');
                if (p.rr_api === 'plan' && ctx.request.method === 'GET') {
                    validatePlan(plan);
                    jsonResponse(ctx, { ok: true, plan, verifyUrl: selfUrl({ rr_verify: key }),
                        armUrl: selfUrl({ rr_api: 'arm', rr_token: key }) });
                } else if (p.rr_api === 'arm' && ctx.request.method === 'POST') {
                    validatePlan(plan);
                    plan.phase = 'submitted';
                    store().put({ key, value: JSON.stringify(plan), ttl: 900 });
                    jsonResponse(ctx, { ok: true });
                } else throw Error('Invalid handoff request.');
            } catch (error) { jsonResponse(ctx, { ok: false, error: error.message }); }
            return;
        }
        const form = ui.createForm({ title: 'Move Commitment to Sales Order' });
        try {
            if (p.rr_verify) {
                const plan = getPlan(String(p.rr_verify));
                const n = native(plan.config.item, plan.config.location);
                const changes = [...plan.sources.filter(r => r.taken > EPS), plan.target];
                const checks = changes.map(r => {
                    const fresh = getOrder(n, r.key);
                    return [r.name + ' / ' + r.orderline, r.after, fresh.committed,
                        Math.abs(fresh.committed - r.after) <= EPS ? 'Verified' : 'Does not match'];
                });
                const verified = checks.every(r => r[3] === 'Verified');
                html(form, 'outcome', '<h2>' + (verified ? 'Commitments verified' : 'Transfer not verified — inspect native commitments') + '</h2>' +
                    breakdown(plan) + table(['Order / line', 'Expected', 'Current', 'Result'], checks));
                log.audit({ title: verified ? 'Native reallocation verified' : 'Native reallocation not verified', details: JSON.stringify({ checks }) });
            } else if (ctx.request.method === 'POST' && p.custpage_action === 'apply') {
                const key = String(p.custpage_token || '');
                const plan = getPlan(key);
                if (plan.phase !== 'preview') throw Error('This preview has already been opened. Reload the Suitelet.');
                validatePlan(plan);
                plan.phase = 'ready';
                store().put({ key, value: JSON.stringify(plan), ttl: 900 });
                const data = JSON.stringify({ token: key, endpoint: selfUrl({ rr_api: 'plan', rr_token: key }) }).replace(/</g, '\\u003c');
                html(form, 'native_handoff', '<h2>Apply through native Reallocate Items</h2>' + breakdown(plan) +
                    '<p>Open the native page to apply the reviewed allocation. Allow the NetSuite popup. Native validation prompts remain visible.</p>' +
                    '<button id="rr_open_native" type="button" style="padding:10px 18px">Open Native Page and Apply</button>' +
                    '<script>(function(){const HANDOFF_DATA=' + data + ';' + HANDOFF_JS + '})();</script>');
            } else {
                const c = config(p);
                if (ctx.request.method === 'POST' && p.custpage_action === 'preview') {
                    const ids = []; const count = ctx.request.getLineCount({ group: 'custpage_sources' });
                    for (let line = 0; line < count; line++) {
                        if (ctx.request.getSublistValue({ group: 'custpage_sources', name: 'custpage_take', line }) === 'T') {
                            ids.push(String(ctx.request.getSublistValue({ group: 'custpage_sources', name: 'custpage_resid', line })));
                        }
                    }
                    const plan = createPlan(c, String(p.custpage_target || ''), ids, String(p.custpage_qty || '').trim());
                    delete plan.native;
                    plan.owner = owner();
                    plan.phase = 'preview';
                    plan.expires = Date.now() + 15 * 60 * 1000;
                    const key = random.generateUUID();
                    store().put({ key, value: JSON.stringify(plan), ttl: 900 });
                    hidden(form, 'action', 'apply'); hidden(form, 'token', key);
                    html(form, 'preview', '<h2>Review proposed transfer</h2>' + breakdown(plan) +
                        '<p>Apply prepares a native-page handoff. On the next screen, Open Native Page and Apply opens a browser window, fills these quantities and clicks its Submit button. Native validation prompts remain visible. This plan expires in fifteen minutes.</p>');
                    form.addSubmitButton({ label: 'Apply Allocation' });
                } else {
                    hidden(form, 'so', c.so);
                    input(form, 'item', 'Item', ui.FieldType.SELECT, c.item, 'item').isMandatory = true;
                    input(form, 'location', 'Location', ui.FieldType.SELECT, c.location, 'location').isMandatory = true;
                    if (!c.item || !c.location) {
                        hidden(form, 'action', 'load');
                        html(form, 'intro', '<p>Select an item and location. Choose the destination sales order line and the orders to take commitment from.</p>');
                        form.addSubmitButton({ label: 'Load Orders' });
                    } else {
                        const n = native(c.item, c.location);
                        const details = orderRows(n);
                        const target = input(form, 'target', 'Destination sales order line', ui.FieldType.SELECT);
                        target.addSelectOption({ value: '', text: 'Select destination sales order line' }); target.isMandatory = true;
                        const targets = details.filter(e => e.isSO && (!c.so || e.id === c.so) && e.remaining > e.committed + EPS);
                        targets.forEach(e => target.addSelectOption({ value: e.key, text: e.name + ' / line ' + e.orderline + ' — needs ' + round(e.remaining - e.committed) }));
                        if (targets.length === 1) target.defaultValue = targets[0].key;
                        if (!targets.length) throw Error('No destination sales order line with a shortfall was returned by NetSuite for this item/location.');
                        input(form, 'qty', 'Maximum quantity to move (blank = SO line shortfall)', ui.FieldType.FLOAT).setHelpText({ help: 'In native reallocation units. Selected sources are consumed in the displayed order.' });
                        const list = form.addSublist({ id: 'custpage_sources', type: ui.SublistType.LIST, label: 'Select orders to take commitment from' });
                        list.addField({ id: 'custpage_take', type: ui.FieldType.CHECKBOX, label: 'Take from' });
                        ['resid', 'reservation', 'type', 'orderline', 'channel', 'committed', 'picked'].forEach((id, i) => {
                            const f = list.addField({ id: PREFIX + id, type: ui.FieldType.TEXT, label: ['Row identity', 'Order', 'Type', 'Order line', 'Channel', 'Committed', 'Picked'][i] });
                            if (id === 'resid') f.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });
                        });
                        const candidates = details.filter(r => r.committed > EPS && (!c.so || r.id !== c.so));
                        candidates.forEach((r, line) => {
                            const data = { resid: r.key, reservation: r.name, type: r.type || '(unknown)', orderline: r.orderline || '-', channel: r.channelText || '(none)', committed: r.committed, picked: r.picked };
                            Object.keys(data).forEach(id => list.setSublistValue({ id: PREFIX + id, line, value: String(data[id]) }));
                        });
                        hidden(form, 'action', 'preview');
                        html(form, 'instructions', '<p>Select the orders to take commitment from. Any source order type exposed by this native grid may be selected, including reservations. ' +
                            'Picked quantities, special-order supply and reservation supply already allocated to demand are excluded. ' +
                            'Quantities use ' + esc(n.units || 'native reallocation units') + '. The next screen shows exactly what will move directly to the destination SO line.</p>');
                        form.addSubmitButton({ label: 'Calculate Allocation' });
                    }
                }
            }
        } catch (error) {
            log.error({ title: 'Reservation reallocation', details: error });
            html(form, 'error', '<h2>Allocation stopped</h2><pre style="white-space:pre-wrap">' + esc(error.name + ': ' + error.message) + '</pre><p>Open the Suitelet again to reload. If a save was attempted, inspect current commitments before retrying.</p>');
        }
        ctx.response.writePage(form);
    }
    return { onRequest };
});
