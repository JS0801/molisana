/**
 * @NApiVersion 2.1
 * @NScriptType ClientScript
 * @NModuleScope SameAccount
 */
define(['N/currentRecord', 'N/url', 'N/https', 'N/ui/dialog'], (currentRecord, url, https, dialog) => {
    // Match these IDs to the Suitelet script and deployment records.
    const SCRIPT_ID = 'customscript_mi_so_inventory_detail_sl';
    const DEPLOYMENT_ID = 'customdeploy_mi_so_inventory_detail_sl';
    let running = false;
    function pageInit() {}
    async function assignInventoryDetail() {
        if (running) return;
        running = true;
        let saved = false;
        try {
            const soid = currentRecord.get().id;
            if (!soid) throw Error('Save the sales order first.');
            const response = await https.post.promise({
                url: url.resolveScript({ scriptId: SCRIPT_ID, deploymentId: DEPLOYMENT_ID, returnExternalUrl: false }),
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ soid })
            });
            let result;
            try { result = JSON.parse(response.body); }
            catch (_) { throw Error('The Suitelet did not return JSON. Check deployment access and your login; refresh the SO to check its inventory detail before retrying.'); }
            if (Number(response.code) !== 200 || !result.success) throw Error(result.message || 'Inventory detail processing failed.');
            saved = Number(result.updated) > 0;
            await dialog.alert({ title: saved ? 'Inventory Detail Updated' : 'No Changes', message: result.message });
        } catch (error) {
            await dialog.alert({ title: 'Inventory Detail', message: error.message || String(error) });
        } finally {
            running = false;
            if (saved) window.location.reload();
        }
    }
    return { pageInit, assignInventoryDetail };
});
