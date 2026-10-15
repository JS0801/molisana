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