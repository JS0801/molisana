/**
 * @NApiVersion 2.1
 * @NScriptType ClientScript
 * @NModuleScope SameAccount
 */
define(['N/currentRecord', 'N/url', 'N/https', 'N/ui/dialog', 'N/ui/message'], (currentRecord, url, https, dialog, message) => {
    // Match these IDs to the Suitelet script and deployment records.
    const SCRIPT_ID = 'customscript_mi_sl_populate_inv_details_';
    const DEPLOYMENT_ID = 'customdeploy_mi_sl_populate_inv_details_';
    let running = false;
    function pageInit() {}
    async function assignInventoryDetail() {
        if (running) return;
        running = true;
        let saved = false;
        let loadingMessage;
        function hideLoading() {
            if (loadingMessage) {
                loadingMessage.hide();
                loadingMessage = null;
            }
        }
        try {
            const soid = currentRecord.get().id;
            if (!soid) throw Error('Save the sales order first.');
            loadingMessage = message.create({
                title: 'Processing Inventory Detail...',
                message: 'Please wait while inventory detail is assigned to this sales order. Keep this page open until processing finishes.',
                type: message.Type.INFORMATION
            });
            loadingMessage.show({ duration: 0 });
            const response = await https.post.promise({
                url: url.resolveScript({ scriptId: SCRIPT_ID, deploymentId: DEPLOYMENT_ID, returnExternalUrl: false }),
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ soid })
            });
            hideLoading();
            let result;
            try { result = JSON.parse(response.body); }
            catch (_) { throw Error('The Suitelet did not return JSON. Check deployment access and your login; refresh the SO to check its inventory detail before retrying.'); }
            if (Number(response.code) !== 200 || !result.success) throw Error(result.message || 'Inventory detail processing failed.');
            saved = Number(result.updated) > 0;
            await dialog.alert({ title: saved ? 'Inventory Detail Updated' : 'No Changes', message: result.message });
        } catch (error) {
            hideLoading();
            await dialog.alert({ title: 'Inventory Detail', message: error.message || String(error) });
        } finally {
            running = false;
            if (saved) window.location.reload();
        }
    }
    return { pageInit, assignInventoryDetail };
});
