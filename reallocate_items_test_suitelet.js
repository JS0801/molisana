/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * Deploy in sandbox, keep Available Without Login unchecked.
 * Diagnostic only: no record.save() and no persistent commitment changes.
 */
define(['N/record', 'N/ui/serverWidget', 'N/runtime', 'N/log'],
    (record, ui, runtime, log) => {
        const escapeHtml = value => String(value == null ? '' : value)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
        const display = value => typeof value === 'object' ? JSON.stringify(value) : String(value == null ? '' : value);

        function onRequest(context) {
            const posted = context.request.method === 'POST';
            const p = context.request.parameters;
            const form = ui.createForm({ title: 'Reallocate Items — Sandbox Script Test' });
            const intro = form.addField({ id: 'custpage_intro', type: ui.FieldType.INLINEHTML, label: 'Instructions' });
            intro.defaultValue = '<p>Select the same item and location used on Reallocate Items, then run the test. ' +
                'This checks unsupported dynamic record creation and sourcing. <b>No changes are saved.</b></p>';
            const item = form.addField({ id: 'custpage_item', type: ui.FieldType.SELECT, source: 'item', label: 'Item' });
            const location = form.addField({ id: 'custpage_location', type: ui.FieldType.SELECT, source: 'location', label: 'Location' });
            item.isMandatory = true;
            location.isMandatory = true;
            const probeLine = form.addField({ id: 'custpage_probe_line', type: ui.FieldType.INTEGER, label: 'Optional test row (zero-based)' });
            probeLine.setHelpText({ help: 'Leave blank for the first run. Use a row number from the results to test editing its quantity in memory only.' });
            const probeQty = form.addField({ id: 'custpage_probe_qty', type: ui.FieldType.FLOAT, label: 'Optional test committed quantity' });
            probeQty.setHelpText({ help: 'Used only with a test row. No transfer is saved.' });
            [item, location, probeLine, probeQty].forEach(field => {
                if (posted && p[field.id] != null) field.defaultValue = String(p[field.id]);
            });
            form.addSubmitButton({ label: 'Run Diagnostic Test' });

            if (posted) {
                let stage = 'validate inputs';
                const results = [];
                const report = (title, details) => {
                    results.push('<h3>' + escapeHtml(title) + '</h3><pre style="white-space:pre-wrap">' +
                        escapeHtml(typeof details === 'string' ? details : JSON.stringify(details, null, 2)) + '</pre>');
                    log.audit({ title, details: typeof details === 'string' ? details : JSON.stringify(details) });
                };
                try {
                    // if (runtime.envType !== runtime.EnvType.SANDBOX) throw new Error('Run this test in SANDBOX only.');
                    // if (!/^\d+$/.test(p.custpage_item || '') || !/^\d+$/.test(p.custpage_location || '')) {
                    //     throw new Error('Select an item and location.');
                    // }
                    stage = 'create dynamic reallocateitem';
                    const rec = record.create({ type: 'reallocateitem', isDynamic: true });
                    report('CREATE PASSED', 'Dynamic reallocateitem created.');
                    stage = 'source item';
                    rec.setValue({ fieldId: 'item', value: Number(p.custpage_item) });
                    stage = 'source location';
                    rec.setValue({ fieldId: 'location', value: Number(p.custpage_location) });
                    const body = {};
                    ['item', 'location', 'quantityonhand', 'quantitycommitted', 'quantityrequired', 'quantitypicked', 'units']
                        .forEach(fieldId => { body[fieldId] = rec.getValue({ fieldId }); });
                    report('SOURCING RESULT', body);
                    stage = 'read order sublist';
                    const fields = rec.getSublistFields({ sublistId: 'order' });
                    const count = rec.getLineCount({ sublistId: 'order' });
                    report('ORDER FIELDS', fields);
                    report('ORDER ROW COUNT', count);
                    results.push('<p>Showing up to 100 rows. Row numbers are zero-based.</p><div style="overflow:auto;max-height:600px">' +
                        '<table border="1" cellpadding="6" style="border-collapse:collapse"><thead><tr><th>Row</th>' +
                        fields.map(f => '<th>' + escapeHtml(f) + '</th>').join('') + '</tr></thead><tbody>');
                    for (let line = 0; line < Math.min(count, 100); line++) {
                        results.push('<tr><td>' + line + '</td>');
                        for (const fieldId of fields) {
                            let value;
                            try { value = rec.getSublistValue({ sublistId: 'order', fieldId, line }); }
                            catch (error) { value = 'READ ERROR: ' + error.message; }
                            results.push('<td>' + escapeHtml(display(value)) + '</td>');
                        }
                        results.push('</tr>');
                    }
                    results.push('</tbody></table></div>');
                    const hasLine = String(p.custpage_probe_line || '').trim() !== '';
                    const hasQty = String(p.custpage_probe_qty || '').trim() !== '';
                    if (hasLine || hasQty) {
                        stage = 'validate optional edit';
                        const line = Number(p.custpage_probe_line);
                        const qty = Number(p.custpage_probe_qty);
                        if (!hasLine || !hasQty || !Number.isInteger(line) || line < 0 || line >= count || !Number.isFinite(qty) || qty < 0) {
                            throw new Error('Provide both a valid row number and a non-negative quantity, or leave both blank.');
                        }
                        const before = rec.getSublistValue({ sublistId: 'order', fieldId: 'quantitycommitted', line });
                        stage = 'edit quantity in memory';
                        rec.selectLine({ sublistId: 'order', line });
                        rec.setCurrentSublistValue({ sublistId: 'order', fieldId: 'quantitycommitted', value: qty });
                        rec.commitLine({ sublistId: 'order' });
                        const after = rec.getSublistValue({ sublistId: 'order', fieldId: 'quantitycommitted', line });
                        report('IN-MEMORY EDIT RESULT', { line, before, requested: qty, after });
                    }
                    report('TEST COMPLETE — NO SAVE', 'No commitments were submitted. A successful result does not yet prove a saved transfer works.');
                } catch (error) {
                    report('TEST FAILED: ' + stage, { name: error.name, message: error.message, stack: error.stack });
                    log.error({ title: 'Reallocation test failed: ' + stage, details: error });
                }
                const output = form.addField({ id: 'custpage_results', type: ui.FieldType.INLINEHTML, label: 'Results' });
                output.defaultValue = results.join('');
            }
            context.response.writePage(form);
        }
        return { onRequest };
    });
