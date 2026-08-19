/**
* @NApiVersion 2.1
* @NScriptType workflowactionscript
*/
define(['N/record', 'N/search', 'N/log'], function (record, search, log) {

    function onAction(scriptContext) {
        let waveId = null;

        try {
            const poRec = scriptContext.newRecord;
            const recID = poRec.id;
            const locationId = poRec.getValue('custbody_tc_shipping_loc');

            if (!recID || !locationId) {
                log.error('Missing Required Fields', {
                    recordId: recID,
                    locationId: locationId
                });
                return;
            }

            // === Wave for Sales Orders ===
            const waveIdSO = createWaveForType({
                recID: recID,
                locationId: locationId,
                tranType: 'SalesOrd',               // transaction.type filter
                waveType: 'SalesOrd',               // wavetype on wave
                searchIdPrefix: 'customsearch17061_wms_nswave_',
                searchTitlePrefix: 'Wave Creation SO '
            });

            if (waveIdSO) {
                waveId = waveIdSO;
            }

            // === Wave for Transfer Orders ===
            const waveIdTO = createWaveForType({
                recID: recID,
                locationId: locationId,
                tranType: 'TrnfrOrd',
                waveType: 'TrnfrOrd',
                searchIdPrefix: 'customsearch_wms_nswave_to_',
                searchTitlePrefix: 'Wave Creation TO '
            });

            if (!waveId && waveIdTO) {
                waveId = waveIdTO;
            }

        } catch (e) {
            log.error("Wave Creation Error (Overall)", e);
        }

        return waveId;
    }

    /**
     * Reusable function:
     * - Builds a transaction search for given type (SalesOrd / TrnfrOrd)
     * - Saves it as a temp saved search
     * - Creates a Wave with that search as template
     * - Deletes the temp saved search
     */
    function createWaveForType(options) {
        const recID           = options.recID;
        const locationId      = options.locationId;
        const tranType        = options.tranType;        // 'SalesOrd' or 'TrnfrOrd'
        const waveType        = options.waveType;        // 'SalesOrd' or 'TrnfrOrd'
        const searchIdPrefix  = options.searchIdPrefix;
        const searchTitlePref = options.searchTitlePrefix;

        try {
            // ---------- Build ONE common-style search ----------
            const transSearch = search.create({
                type: "transaction",
                filters: [
                    ["type", "anyof", tranType],
                    "AND",
                    ["custcol_tc_scheduled_ship_date", "isnotempty", ""],
                    "AND",
                    ["inventorylocation", "noneof", "@NONE@"],
                    "AND",
                    ["custcol_tc_related_shipping_record", "anyof", recID],
                    "AND",
                    [["quantitycommitted", "greaterthan", "0"], "OR", ["item.custitem_tc_is_crating_item","is","T"]]
                ],
                columns: [
                    "tranid",
                    search.createColumn({ name: "item", sort: search.Sort.ASC }),
                    "memo",
                    "unpickedorderquantity",
                    "unit",
                    "trandate",
                    "entity",
                    "memomain",
                    "transferlocation",
                    search.createColumn({
                        name: "formulatext",
                        formula: "{otherrefnum}",
                        label: "PO NUMBER"
                    })
                ]
            });

            const count = transSearch.runPaged().count;
            log.debug('search count (' + tranType + ')', count);

            if (count <= 0) {
                log.audit(
                    "No Orders Found (" + tranType + ")",
                    "Skipping wave creation for CSR: " + recID
                );
                return null;
            }

            const token = generateToken();
            transSearch.id    = searchIdPrefix + token;
            transSearch.title = searchTitlePref + token;

            log.debug('Token (' + tranType + ')', token);

            const savedSearchId = transSearch.save();
            log.debug('savedSearchId (' + tranType + ')', savedSearchId);

            // ---------- Create Wave ----------
            const wave = record.create({
                type: record.Type.WAVE,
                isDynamic: true
            });

            wave.setValue('location', locationId);
            wave.setValue('custbody_ds_related_ship_rec', recID);
            wave.setValue('wavetype', waveType);           // differs by type
            wave.setValue('priority', '1');
            wave.setValue('picktype', 'MULTI');
            wave.setValue('newwavestatus', 'PENDING');
            wave.setValue('searchtemplateid', savedSearchId);
            wave.setValue('custbody_tc_related_csr', recID);

            const waveId = wave.save();
            log.debug("Wave Created (" + tranType + ")", waveId);

            // Clean up temp search
            try {
                search.delete({ id: savedSearchId });
            } catch (delErr) {
                log.error("Failed to delete temp search (" + tranType + ")", delErr);
            }

            return waveId;

        } catch (e) {
            log.error("Wave Creation Error (" + tranType + ")", e);
            return null;
        }
    }

    function generateToken() {
        const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
        let token = '';
        for (let i = 0; i < 10; i++) {
            token += chars[Math.floor(Math.random() * chars.length)];
        }
        return token;
    }

    return { onAction: onAction };

});
