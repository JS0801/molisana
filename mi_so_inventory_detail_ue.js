/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 * @NModuleScope SameAccount
 */
define(['N/record'], (record) => {
    function beforeLoad(context) {
        if (context.type !== context.UserEventType.VIEW ||
            context.newRecord.type !== record.Type.SALES_ORDER || !context.newRecord.id) return;
        context.form.clientScriptModulePath = './mi_so_inventory_detail_cs.js';
        context.form.addButton({
            id: 'custpage_mi_assign_inventory',
            label: 'Assign Inventory Detail',
            functionName: 'assignInventoryDetail'
        });
    }
    return { beforeLoad };
});
