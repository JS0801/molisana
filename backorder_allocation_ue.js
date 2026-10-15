/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 * @NModuleScope SameAccount
 * Optional Sales Order button. Put the companion client file in the same folder.
 * Required script parameters (Free-Form Text):
 * custscript_rr_suitelet_script = Suitelet script ID
 * custscript_rr_suitelet_deploy = Suitelet deployment ID
 */
define(['N/ui/serverWidget', 'N/runtime', 'N/url'], (ui, runtime, url) => {
    function beforeLoad(ctx) {
        if (ctx.type !== ctx.UserEventType.VIEW || runtime.executionContext !== runtime.ContextType.USER_INTERFACE) return;
        const script = runtime.getCurrentScript();
        const scriptId = 'customscript_sl_allocate_qty';
        const deploymentId = 'customdeploy_sl_allocate_qty';
        if (!scriptId || !deploymentId) return;
        const target = url.resolveScript({ scriptId, deploymentId, params: { soid: ctx.newRecord.id } });
        const f = ctx.form.addField({ id: 'custpage_rr_url', label: 'Reservation allocation URL', type: ui.FieldType.LONGTEXT });
        f.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });
        f.defaultValue = target;
        ctx.form.clientScriptModulePath = './backorder_allocation_modal_client_v2.js';
        ctx.form.addButton({ id: 'custpage_rr_allocate', label: 'Allocate Backordered Items', functionName: 'openBackorderAllocationModalV2' });
    }
    return { beforeLoad };
});
