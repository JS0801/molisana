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
    const store = () => cache.getCache({ name: 'reservation_reallocation_plans_v1', scope: cache.Scope.PRIVATE });
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
                if (identity.includes(fieldId)) {
                    try { labels[fieldId] = rec.getSublistText({ sublistId: 'order', fieldId, line }); } catch (_) { /* value is sufficient */ }
                }
            });
            rows.push({ line, values, labels });
        }
        return { rec, fields, identity, rows, units: text(rec, 'units') || String(val(rec, 'units') || '') };
    }
    function reservations(item, location) {
        const list = search.create({ type: 'orderreservation', filters: [
            ['mainline', 'is', 'T'], 'AND', ['item', 'anyof', item], 'AND', ['location', 'anyof', location]
        ], columns: ['internalid', 'tranid'] }).run().getRange({ start: 0, end: 1000 });
        if (list.length >= 1000) throw Error('Reservation search reached its limit; narrow the scope.');
        const seen = new Set();
        return list.filter(r => { const id = String(r.id); if (seen.has(id)) return false; seen.add(id); return true; });
    }
    function match(n, id, tranid) {
        const matches = n.rows.filter(row => n.identity.some(f => {
            const v = String(row.values[f] == null ? '' : row.values[f]);
            const t = String(row.labels[f] == null ? '' : row.labels[f]);
            // Numeric IDs only on identity fields; exact transaction numbers
            // can also be exposed as the field's display text.
            return v === String(id) || (tranid && (v === tranid || t === tranid));
        }));
        if (matches.length !== 1) throw Error('Expected one native row for reservation ' + tranid + ' (' + id + '), found ' + matches.length + '.');
        return matches[0];
    }
    function reservation(id, item, location) {
        const r = record.load({ type: 'orderreservation', id, isDynamic: false });
        if (String(val(r, 'item')) !== String(item) || String(val(r, 'location')) !== String(location)) throw Error('Reservation item/location mismatch.');
        return {
            id: String(id), name: String(val(r, 'tranid') || val(r, 'name') || id),
            channel: String(val(r, 'saleschannel') || ''), channelText: text(r, 'saleschannel'),
            committed: num(val(r, 'quantitycommitted')), demand: num(val(r, 'allocateddemand')),
            closed: yes(val(r, 'closed')), firm: yes(val(r, 'commitmentfirm')),
            units: String(val(r, 'units') || ''), unitsText: text(r, 'units')
        };
    }
    function snapshot(n) {
        // Entire native grid, including unrelated orders, must remain unchanged
        // between preview and apply. This is optimistic validation, not a lock.
        return JSON.stringify(n.rows.map(r => r.values));
    }
    function assertUnits(r, row) {
        if (Math.abs(r.committed - num(row.values.quantitycommitted)) > EPS) {
            throw Error('Reservation ' + r.name + ' and native row committed quantities differ. Reload or check units before transferring.');
        }
    }
    function createPlan(c, targetId, sourceIds, requested) {
        if (!sourceIds.length || sourceIds.length > 20) throw Error('Select between 1 and 20 source reservations.');
        if (new Set(sourceIds).size !== sourceIds.length || sourceIds.includes(targetId)) throw Error('Source reservations must be distinct and different from the destination.');
        const n = native(c.item, c.location);
        const target = reservation(targetId, c.item, c.location);
        if (target.closed || target.firm) throw Error('Destination reservation is closed or firm.');
        if (c.channel && target.channel !== c.channel) throw Error('Destination reservation must match the SO sales channel.');
        const destRow = match(n, target.id, target.name);
        assertUnits(target, destRow);
        const gap = Math.max(0, num(destRow.values.quantityremaining) - num(destRow.values.quantitycommitted));
        const limit = requested === '' ? gap : num(requested);
        if (limit <= 0) throw Error('The requested quantity must be greater than zero.');
        let needed = Math.min(gap, limit);
        const sources = [];
        const usedRows = new Set([destRow.line]);
        // Stable source order: ascending reservation internal ID, shown in preview.
        sourceIds.sort((a, b) => Number(a) - Number(b));
        for (const id of sourceIds) {
            const src = reservation(id, c.item, c.location);
            if (src.closed || src.firm) throw Error('Source ' + src.name + ' is closed or firm.');
            if (src.channel === target.channel) throw Error('Source ' + src.name + ' belongs to the destination channel.');
            if (src.units !== target.units) throw Error('Source and destination must use the same reservation units.');
            const row = match(n, src.id, src.name);
            if (usedRows.has(row.line)) throw Error('Native row identity is ambiguous; refusing to transfer.');
            usedRows.add(row.line);
            assertUnits(src, row);
            const picked = num(row.values.quantitypicked);
            // Conservative: do not take supply already allocated to demand or picked.
            const available = Math.max(0, src.committed - src.demand - picked);
            const taken = round(Math.min(needed, available));
            needed = round(needed - taken);
            sources.push({ ...src, line: row.line, available, taken, after: round(src.committed - taken) });
        }
        const moved = round(sources.reduce((sum, s) => sum + s.taken, 0));
        if (moved <= EPS) throw Error('No releasable commitment in the selected reservations, or the destination has no shortfall.');
        return { config: c, target: { ...target, line: destRow.line, after: round(target.committed + moved) }, sources,
            requested, moved, gap, unmet: round(Math.min(gap, limit) - moved), units: n.units,
            snapshot: snapshot(n), native: n };
    }
    function breakdown(p) {
        return '<p><b>Quantity to move: ' + esc(p.moved) + ' ' + esc(p.units) + '</b>. Unfilled portion of this request: ' + esc(p.unmet) + '.</p>' +
            table(['Source reservation', 'Source channel', 'Before', 'Taken', 'After', 'Destination reservation', 'Destination channel'],
                p.sources.map(s => [s.name, s.channelText || s.channel, s.committed, s.taken, s.after, p.target.name, p.target.channelText || p.target.channel])) +
            '<p>Destination commitment: <b>' + esc(p.target.committed) + ' → ' + esc(p.target.after) + '</b>.</p>';
    }
    function apply(plan) {
        const current = createPlan(plan.config, plan.target.id, plan.sources.map(s => s.id), plan.requested);
        if (current.snapshot !== plan.snapshot || current.moved !== plan.moved ||
            JSON.stringify(current.sources) !== JSON.stringify(plan.sources) || JSON.stringify(current.target) !== JSON.stringify(plan.target)) {
            throw Error('Allocations or reservation settings changed since preview. Reload and preview again.');
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
            for (const r of changes) {
                const fresh = reservation(r.id, plan.config.item, plan.config.location);
                const ok = Math.abs(fresh.committed - r.after) <= EPS;
                if (!ok) verified = false;
                checks.push([r.name, r.after, fresh.committed, ok ? 'Verified' : 'Does not match']);
            }
        } catch (error) { verified = false; checks.push(['Verification error', '', '', error.message]); }
        log.audit({ title: 'Reservation reallocation saved', details: JSON.stringify({ user: owner(), item: plan.config.item, location: plan.config.location, savedId, moved: plan.moved, target: plan.target.id, sources: plan.sources.map(s => ({ id: s.id, taken: s.taken })), verified }) });
        return { savedId, verified, checks };
    }
    function config(p) {
        const c = { item: String(p.custpage_item || ''), location: String(p.custpage_location || ''), so: String(p.soid || p.custpage_so || ''), channel: '' };
        if (c.so) {
            const so = record.load({ type: 'salesorder', id: c.so });
            c.channel = String(val(so, 'saleschannel') || '');
            if (!c.channel) throw Error('This SO has no sales channel.');
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
        const form = ui.createForm({ title: 'Move Commitment Between Order Reservations' });
        try {
      //      if (runtime.envType !== runtime.EnvType.SANDBOX) throw Error('This initial operational version is restricted to sandbox until verified.');
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
                    (result.verified ? '' : '<p>Do not repeat the transfer until you inspect the native reservation quantities.</p>') +
                    breakdown(plan) + table(['Reservation', 'Expected commitment', 'Read-back commitment', 'Result'], result.checks));
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
                        html(form, 'intro', '<p>Select an item and location to find its order reservations. Quantities move between reservations, not directly to the SO.</p>');
                        form.addSubmitButton({ label: 'Load Reservations' });
                    } else {
                        const n = native(c.item, c.location);
                        const found = reservations(c.item, c.location);
                        const eligible = []; const skipped = [];
                        for (const r of found) {
                            const id = String(r.id), name = String(r.getValue('tranid') || '');
                            try { const row = match(n, id, name); eligible.push({ id, name, row }); }
                            catch (error) { skipped.push(name + ': ' + error.message); }
                        }
                        if (!eligible.length) throw Error('No reservations could be mapped uniquely to native rows. ' + skipped.join(' ') + ' Native fields: ' + n.fields.join(', '));
                        // Load at most 40 reservations for channel and protection details.
                        if (eligible.length > 40) throw Error('More than 40 reservations; narrow the scope before proceeding.');
                        const details = eligible.map(e => ({ ...e, info: reservation(e.id, c.item, c.location) }));
                        const target = input(form, 'target', 'Destination order reservation', ui.FieldType.SELECT);
                        target.addSelectOption({ value: '', text: 'Select destination reservation' }); target.isMandatory = true;
                        const targets = details.filter(e => !e.info.closed && !e.info.firm && (!c.channel || e.info.channel === c.channel) && num(e.row.values.quantityremaining) > num(e.row.values.quantitycommitted) + EPS);
                        targets.forEach(e => target.addSelectOption({ value: e.id, text: e.name + ' — ' + (e.info.channelText || e.info.channel) }));
                        if (targets.length === 1) target.defaultValue = targets[0].id;
                        if (!targets.length) throw Error('No open, non-firm destination reservation with a shortfall for this selection/channel.');
                        input(form, 'qty', 'Maximum quantity to move (blank = destination shortfall)', ui.FieldType.FLOAT).setHelpText({ help: 'In the displayed native reallocation units. Sources are consumed in ascending reservation ID order.' });
                        const list = form.addSublist({ id: 'custpage_sources', type: ui.SublistType.LIST, label: 'Select source reservations' });
                        list.addField({ id: 'custpage_take', type: ui.FieldType.CHECKBOX, label: 'Take from' });
                        ['resid', 'reservation', 'channel', 'committed', 'demand', 'protection'].forEach((id, i) => list.addField({ id: PREFIX + id, type: ui.FieldType.TEXT, label: ['Internal ID', 'Reservation', 'Channel', 'Committed', 'Allocated demand', 'Protection'][i] }));
                        details.forEach((e, line) => {
                            const r = e.info;
                            const data = { resid: r.id, reservation: r.name, channel: r.channelText || r.channel || '(none)', committed: r.committed, demand: r.demand, protection: r.closed ? 'Closed — excluded' : r.firm ? 'Firm — excluded' : 'Open' };
                            Object.keys(data).forEach(id => list.setSublistValue({ id: PREFIX + id, line, value: String(data[id]) }));
                        });
                        hidden(form, 'action', 'preview');
                        html(form, 'instructions', '<p>Select sources in other channels. Closed/firm reservations and demand-backed quantities are protected. ' +
                            'Quantities use ' + esc(n.units || 'native reallocation units') + '. The next screen shows exactly what will move.</p>' +
                            (skipped.length ? '<details><summary>Reservations unavailable on this native grid</summary><pre>' + esc(skipped.join('\n')) + '</pre></details>' : ''));
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
