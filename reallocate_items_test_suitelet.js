/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * Unsupported reallocateitem adapter. Requires sandbox validation before production.
 * No HTTP/page scraping. Saves one native reallocation record per operation.
 */
define(['N/record', 'N/search', 'N/ui/serverWidget', 'N/runtime', 'N/cache', 'N/crypto/random', 'N/log'],
(record, search, ui, runtime, cache, random, log) => {
    const PREFIX = 'custpage_';
    const EPS = 0.000001;
    const esc = v => String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    const num = v => { const n = Number(v || 0); if (!Number.isFinite(n)) throw Error('Invalid quantity: ' + v); return n; };
    const round = n => Math.round(n * 1000000) / 1000000;
    const yes = v => v === true || v === 'T';
    const store = () => cache.getCache({ name: 'order_reallocation_plans_v2', scope: cache.Scope.PRIVATE });
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
    function apply(plan) {
        const current = createPlan(plan.config, plan.target.key, plan.sources.map(s => s.key), plan.requested);
        if (current.snapshot !== plan.snapshot || current.moved !== plan.moved ||
            JSON.stringify(current.sources) !== JSON.stringify(plan.sources) || JSON.stringify(current.target) !== JSON.stringify(plan.target)) {
            throw Error('Allocations or order settings changed since preview. Reload and preview again.');
        }
        const n = current.native;
        const changes = [...current.sources.filter(s => s.taken > EPS), current.target];
        function setQuantity(line, quantity) {
            n.rec.selectLine({ sublistId: 'order', line });
            n.rec.setCurrentSublistValue({ sublistId: 'order', fieldId: 'commit', value: quantity > EPS });
            n.rec.setCurrentSublistValue({ sublistId: 'order', fieldId: 'quantitycommitted', value: quantity });
            n.rec.commitLine({ sublistId: 'order' });
        }
        // Release source quantities first, then increase the destination, all
        // in one native record. Never save separate partial source changes.
        changes.forEach(r => setQuantity(r.line, r.after));
        for (const row of n.rows) {
            const change = changes.find(r => r.line === row.line);
            const expected = change ? change.after : num(row.values.quantitycommitted);
            const actual = num(n.rec.getSublistValue({ sublistId: 'order', fieldId: 'quantitycommitted', line: row.line }));
            if (Math.abs(expected - actual) > EPS) throw Error('NetSuite changed an unexpected quantity on native row ' + row.line + '. Nothing saved.');
        }
        let savedId;
        try { savedId = n.rec.save({ enableSourcing: true, ignoreMandatoryFields: false }); }
        catch (error) { throw Error('Save did not return success. Inspect native commitments before retrying. NetSuite: ' + error.name + ': ' + error.message); }
        const checks = [];
        let verified = true;
        try {
            const reread = native(plan.config.item, plan.config.location);
            for (const r of changes) {
                const fresh = getOrder(reread, r.key);
                const ok = Math.abs(fresh.committed - r.after) <= EPS;
                if (!ok) verified = false;
                checks.push([r.name + ' / ' + r.orderline, r.after, fresh.committed, ok ? 'Verified' : 'Does not match']);
            }
        } catch (error) { verified = false; checks.push(['Verification error', '', '', error.message]); }
        log.audit({ title: 'Reservation reallocation saved', details: JSON.stringify({ user: owner(), item: plan.config.item, location: plan.config.location, savedId, moved: plan.moved, target: plan.target.key, sources: plan.sources.map(s => ({ key: s.key, taken: s.taken })), verified }) });
        return { savedId, verified, checks };
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
        const form = ui.createForm({ title: 'Move Commitment to Sales Order' });
        try {
          //  if (runtime.envType !== runtime.EnvType.SANDBOX) throw Error('This initial operational version is restricted to sandbox until verified.');
            if (ctx.request.method === 'POST' && p.custpage_action === 'apply') {
                const key = String(p.custpage_token || '');
                const raw = store().get({ key });
                if (!raw) throw Error('Preview expired or was already used. Open the Suitelet again to reload.');
                const plan = JSON.parse(raw);
                if (plan.owner !== owner()) throw Error('This preview belongs to another user or role.');
                // Consume before save; refresh will not replay this plan.
                // N/cache is not an atomic distributed lock. Do not run concurrent submissions.
                store().remove({ key });
                const result = apply(plan);
                html(form, 'outcome', '<h2>' + (result.verified ? 'Transfer saved and verified' : 'Save returned — verification needs review') + '</h2>' +
                    (result.verified ? '' : '<p>Do not repeat the transfer until you inspect the native order quantities.</p>') +
                    breakdown(plan) + table(['Order / line', 'Expected commitment', 'Read-back commitment', 'Result'], result.checks));
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
                    const key = random.generateUUID();
                    store().put({ key, value: JSON.stringify(plan), ttl: 300 });
                    hidden(form, 'action', 'apply'); hidden(form, 'token', key);
                    html(form, 'preview', '<h2>Review proposed transfer</h2>' + breakdown(plan) +
                        '<p>Apply saves these quantities to NetSuite. This preview expires in five minutes. Submit once and wait for the result.</p>');
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
