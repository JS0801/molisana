/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 *
 * TOSAN SHIPPING PROVIDER – Weekly Shipping Schedule
 *
 * Logic:
 * - Filter Sales Orders by Shipping Provider.
 * - Expected Shipping Date drives week/day allocation.
 * - Shows Sales Orders for current calendar year.
 * - Groups by Week -> Zone -> Day.
 * - Weekends are skipped.
 * - UI is loaded from separate HTML file in File Cabinet.
 * - Filter / Save / Email / Upload uses AJAX, no full page refresh.
 *
 * Proof of Delivery (POD):
 * - Replaces the old single POD Link URL field.
 * - Users upload one or more files/photos per Sales Order.
 * - Files are created in the File Cabinet and attached to the SO so they
 *   appear under Communication > Files on the Sales Order record.
 * - The grid lists the currently attached POD files per SO (with links).
 *
 * Required Script Parameter:
 * custscript_tosan_html_file_id = File Cabinet internal ID or path of tosan_shipping_provider.html
 *
 * Optional Script Parameter:
 * custscript_tosan_pod_folder_id = File Cabinet folder internal ID where POD files are stored.
 */
define([
    'N/search',
    'N/record',
    'N/email',
    'N/runtime',
    'N/format',
    'N/file',
    'N/log',
    'N/url',
    'N/encode'
], function (
    search,
    record,
    email,
    runtime,
    format,
    file,
    log,
    url,
    encode
) {

    /* ============================= CONFIG ============================= */

    var PROVIDER_LIST = 'customlist_mi_shipping_provider_list';
    var PROVIDER_FIELD = 'custbody_shipping_provider_transaction';
    var ZONE_FIELD = 'custbody_mi_delivery_zone';
    var DATE_FIELD = 'custbody_expected_shipping_date';
    var AVAILABLE_PICKUP_AFTER_FIELD = 'custbody_available_to_pick_up_after';
    var ACTUAL_PICKUP_DATE_FIELD = 'custbody_actual_pick_up_date';

    // UI Label = Pallet
    var VOLUME_FIELD = 'custbody_total_so_volume';
    var TRIP_LINE_CUSTOM_QTY_FIELD = 'custcol_mi_quantity';

    // UI Label = Cases
    var CASES_FIELD = 'custbody_total_cases_to_ship';

    // Line-level search source fields (used to compute Pallet / Cases / Weight from lines)
    var LINE_ITEM_VOLUME_FIELD = 'item.custitem_item_volume';
    var LINE_WEIGHT_FIELD = 'custcol_atlas_line_item_weight';

    // shipping address (header-level) fields
    var CITY_FIELD = 'shipcity';
    var ZIP_FIELD = 'shipzip';

    // Extra header fields carried through from the saved search
    var MEMO_FIELD = 'memomain';
    var DELIVERY_NOTE_FIELD = 'custbody_delivery_note';

    // Proof of Delivery upload — files attached to the SO (Communication > Files)
    var POD_UPLOAD_FOLDER_PARAM = 'custscript_tosan_pod_folder_id';
    var DEFAULT_POD_FOLDER_ID = -15; // -15 = SuiteScripts; override with the param above

    var DEFAULT_NAME = 'TOSAN SHIPPING PROVIDER';

    // ---------------------------------------------------------------------
    // PORTAL SCOPING — the SAME script powers multiple portals (URLs) via
    // separate deployments. Each deployment sets ONE parameter:
    //
    //   custscript_shipping_provider = <provider>  -> show ONLY that provider
    //   custscript_shipping_provider = (empty)      -> show ALL providers
    //
    // The value may be the provider's internal id (List/Record param) OR its
    // name (Free-Form Text param) — both are handled.
    // ---------------------------------------------------------------------
    var SHIPPING_PROVIDER_PARAM = 'custscript_shipping_provider_1';

    var PORTAL_TITLE_PARAM = 'custscript_tosan_portal_title';

    var HTML_FILE_PARAM = 'custscript_html_file_id';
    var DEFAULT_HTML_PATH = 'SuiteScripts/TOSAN Create Trip — HTML.html';

    var ZONES = [2, 3, 4, 5, 6, 7, 8, 9, 10];
    var ZONE_LABEL_MAP = {
        '2': 'Zone 2/GTA',
        '3': 'Zone 3/Ottawa',
        '4': 'Zone 4/Montreal',
        '5': 'Zone 5/Oshawa',
        '6': 'Zone 6/Barrie-NewMarket-Brantford',
        '7': 'Zone 7/Windsor',
        '8': 'Zone 8/Hamilton and Niagara Guelph-Georgetown',
        '9': 'Zone 9/London'
    };


    var DAY_SHORT = {
        1: 'MON',
        2: 'TUE',
        3: 'WED',
        4: 'THU',
        5: 'FRI'
    };

    // ---------------------- TRIP BUILDER config ----------------------
    var TRIP_TRANTYPE = 'customsale_trip';
    var TRIP_LINE_SUBLIST = 'item';
    // Sales-type transaction requires an entity (customer). A trip spans many
    // customers, so we use one fixed placeholder customer on the header. The real
    // per-order customer is available via the SO reference (custcol_trip_so).
    var TRIP_CUSTOMER_ID = '7122';
    // Header-level location for the trip transaction.
    var TRIP_LOCATION_ID = '311';

    var LINE_TEMP_FIELD = 'item.custitem8';

    var TRIP_PROVIDER_FIELD = 'custbody_trip_provider';
    var TRIP_TRUCK_FIELD = 'custbody_trip_truck';
    var TRIP_DRIVER_FIELD = 'custbody_trip_driver';
    var TRIP_SHIP_DATE_FIELD = 'custbody_trip_ship_date';
    var TRIP_TOTAL_STOPS_FIELD = 'custbody_trip_total_stops';
    var TRIP_TOTAL_PALLETS_FIELD = 'custbody_trip_total_pallets';
    var TRIP_TOTAL_CASES_FIELD = 'custbody_trip_total_cases';
    var TRIP_TOTAL_WEIGHT_FIELD = 'custbody_trip_total_weight';
    var TRIP_OVERRIDE_TOTALS_FIELD = 'custbody_trip_override_totals';
    var TRIP_TEMP_CONTROL_FIELD = 'custbody_trip_temp_control';
    var TRIP_TEMP_SET_FIELD = 'custbody_trip_temp_set';
    var TRIP_TEMP_CONFLICT_FIELD = 'custbody_trip_temp_conflict';

    var TRIP_LINE_SO_FIELD = 'custcol_trip_so';
    var TRIP_LINE_STOP_SEQ_FIELD = 'custcol_trip_stop_seq';
    var TRIP_LINE_ZONE_FIELD = 'custcol_trip_zone';
    var TRIP_LINE_PALLETS_FIELD = 'custcol_trip_pallets';
    var TRIP_LINE_CASES_FIELD = 'custcol_trip_cases';
    var TRIP_LINE_WEIGHT_FIELD = 'custcol_trip_weight';
    var TRIP_LINE_TEMP_SET_FIELD = 'custcol_trip_temp_set';
    var TRIP_LINE_TEMP_CONTROL_FIELD = 'custcol_trip_temp_control';

    var SO_TRIP_FIELD = 'custbody_trip';
    var SO_RESERVED_BY_FIELD = 'custbody_ttrip_reserved_by';
    var SO_RESERVED_AT_FIELD = 'custbody_trip_reserved_at';
    var SO_LINE_RELATED_TRIP_FIELD = 'custcol_mi_related_trip_record';

    // Trucks are a plain custom LIST (id + name only). A list cannot carry
    // capability fields, so temp-control validation against the truck is not
    // available; the load-side temp warning is advisory only.
    var TRUCK_LIST = 'customlist_truck';

    // Driver custom list
    var DRIVER_LIST = 'customlist_driver';

    var RESERVE_TTL_PARAM = 'custscript_tosan_reserve_ttl_min';
    var DEFAULT_RESERVE_TTL_MIN = 15;

    var TEMP_AMBIENT = 'AMBIENT';
    var TEMP_R = 'R';
    var TEMP_TC = 'TC';
    var TEMP_TCR = 'TC/R';

    var TEMP_COLORS = {
        'AMBIENT': '#639922',
        'R': '#E24B4A',
        'TC': '#185FA5',
        'TC/R': '#534AB7'
    };

    /* ============================ ENTRY ============================ */

    function onRequest(context) {
        var req = context.request;
        var res = context.response;

        var isAjax = req.parameters.tosan_ajax === 'T';
        var action = req.parameters.tosan_action || '';

        if (isAjax) {
            handleAjax(req, res, action);
            return;
        }

        renderPage(res);
    }

    /* ============================ AJAX ============================ */

    function handleAjax(req, res, action) {
        try {
            var providers = getProviders();
            var requested = req.parameters.tosan_provider || '';
            var selectedProvider = resolveProviderInScope(requested, providers);
            var recipient = req.parameters.tosan_recipient || getUserEmail();

            var result;

            if (action === 'data') {
                result = getDataResponse(selectedProvider, providers);
            } else if (action === 'save') {
                result = saveAndRefresh(req.parameters.tosan_payload, selectedProvider, providers);
            } else if (action === 'remove') {
                result = removeAndRefresh(req.parameters.tosan_soid, selectedProvider, providers);
            } else if (action === 'upload') {
                result = uploadAndRefresh(req, selectedProvider, providers);
            } else if (action === 'email') {
                result = emailAndRefresh(selectedProvider, providers, recipient);
            } else if (action === 'trip_unassigned') {
                result = getUnassignedResponse(selectedProvider, providers);
            } else if (action === 'trip_trucks') {
                result = { success: true, trucks: getTrucks() };
            } else if (action === 'trip_drivers') {
                result = { success: true, drivers: getDrivers() };
            } else if (action === 'trip_my_reserved') {
                result = { success: true, myReserved: getMyReserved() };
            } else if (action === 'trip_reserve') {
                result = reserveSo(req.parameters.tosan_soid);
            } else if (action === 'trip_release') {
                result = releaseSo(req.parameters.tosan_soid);
            } else if (action === 'trip_submit') {
                result = submitTrip(req, selectedProvider);
            } else {
                result = {
                    success: false,
                    message: 'Invalid action: ' + action
                };
            }

            writeJson(res, result);

        } catch (e) {
            log.error('TOSAN AJAX Error', e);

            writeJson(res, {
                success: false,
                message: 'Error: ' + (e.message || e)
            });
        }
    }

    function getDataResponse(providerId, providers) {
        var ctx = {};
        var data = getData(providerId, ctx);
        var win = getWeekWindow();

        return {
            success: !ctx.error,
            message: ctx.error || '',
            selectedProvider: providerId,
            providerName: providerNameById(providers, providerId),
            currentWeek: isoWeekInfo(new Date()),
            weekWindow: {
                start: win.startYmd,
                end: win.endYmd,
                startDisplay: win.startStr,
                endDisplay: win.endStr
            },
            zones: ZONES,
            zoneOptions: getZoneOptions(),
            zoneLabelMap: ZONE_LABEL_MAP,
            dayShort: DAY_SHORT,
            tempColors: TEMP_COLORS,
            totals: totals(data),
            data: data
        };
    }

    function saveAndRefresh(payload, providerId, providers) {
        var saveMessage = doSave(payload);
        var refreshed = getDataResponse(providerId, providers);

        refreshed.message = saveMessage + (refreshed.message ? ' | ' + refreshed.message : '');
        refreshed.success = refreshed.success !== false;

        return refreshed;
    }

    function removeAndRefresh(soId, providerId, providers) {
        var removeMessage = doRemove(soId);
        var refreshed = getDataResponse(providerId, providers);

        refreshed.message = removeMessage + (refreshed.message ? ' | ' + refreshed.message : '');
        refreshed.success = refreshed.success !== false;

        return refreshed;
    }

    function uploadAndRefresh(req, providerId, providers) {
        var uploadMessage = doUpload(req);
        var refreshed = getDataResponse(providerId, providers);

        refreshed.message = uploadMessage + (refreshed.message ? ' | ' + refreshed.message : '');
        refreshed.success = refreshed.success !== false;

        return refreshed;
    }

    function emailAndRefresh(providerId, providers, recipient) {
        var emailMessage = doEmail(providerId, providers, recipient);
        var refreshed = getDataResponse(providerId, providers);

        refreshed.message = emailMessage + (refreshed.message ? ' | ' + refreshed.message : '');
        refreshed.success = refreshed.success !== false;

        return refreshed;
    }

    /* ============================ PAGE LOAD ============================ */

    function renderPage(res) {
        var providers = getProviders();
        var selectedProvider = getDefaultProviderId(providers);

        // Heading: explicit title param wins; else the scoped provider's name
        // (single-provider portal); else a generic title (all-providers portal).
        var portalTitle = getPortalTitle();

        if (!portalTitle) {
            if (getConfiguredProvider() && providers.length) {
                portalTitle = providers[0].name + ' Weekly Shipping Schedule';
            } else {
                portalTitle = 'Weekly Shipping Schedule';
            }
        }

        var bootstrap = {
            apiUrl: '',
            providers: providers,
            selectedProvider: selectedProvider,
            selectedProviderName: providerNameById(providers, selectedProvider),
            portalTitle: portalTitle,
            recipient: getUserEmail(),
            currentWeek: isoWeekInfo(new Date()),
            zones: ZONES,
            zoneOptions: getZoneOptions(),
            zoneLabelMap: ZONE_LABEL_MAP,
            dayShort: DAY_SHORT,
            tempColors: TEMP_COLORS
        };

        var html = loadHtmlTemplate();
        html = html.replace(/__TOSAN_BOOTSTRAP_JSON__/g, safeJsonForScript(bootstrap));

        res.setHeader({
            name: 'Content-Type',
            value: 'text/html; charset=utf-8'
        });

        res.write({
            output: html
        });
    }

    function loadHtmlTemplate() {
        var scriptObj = runtime.getCurrentScript();

        var htmlFileId = scriptObj.getParameter({
            name: HTML_FILE_PARAM
        }) || DEFAULT_HTML_PATH;

        try {
            return file.load({
                id: htmlFileId
            }).getContents();

        } catch (e) {
            log.error('HTML template load failed', {
                htmlFileId: htmlFileId,
                error: e
            });

            return '<!doctype html>' +
                '<html>' +
                '<body style="font-family:Arial;padding:24px">' +
                '<h2>TOSAN Shipping Schedule</h2>' +
                '<p style="color:#b91c1c"><b>HTML file could not be loaded.</b></p>' +
                '<p>Upload <code>tosan_shipping_provider.html</code> to File Cabinet and set script parameter ' +
                '<code>' + esc(HTML_FILE_PARAM) + '</code> to the internal ID or path.</p>' +
                '<pre>' + esc(e.message || e) + '</pre>' +
                '</body>' +
                '</html>';
        }
    }

    /* ============================ DATA ============================ */

    // Reads custscript_shipping_provider. Empty -> null (show all).
    function getConfiguredProvider() {
        try {
            var v = runtime.getCurrentScript().getParameter({ name: SHIPPING_PROVIDER_PARAM });

            // List/Record params may return a number or {value:..} style; normalise.
            if (v && typeof v === 'object' && typeof v.value !== 'undefined') {
                v = v.value;
            }

            v = String(v == null ? '' : v).trim();

            return v || null;
        } catch (e) {
            return null;
        }
    }

    // Heading shown in the portal. Explicit title param wins; otherwise the
    // configured provider's name; otherwise a generic title.
    function getPortalTitle() {
        try {
            var v = runtime.getCurrentScript().getParameter({ name: PORTAL_TITLE_PARAM });
            v = String(v || '').trim();

            if (v) {
                return v;
            }
        } catch (e) {
            // fall through
        }

        return ''; // resolved against the scoped provider list in renderPage
    }

    function getProviders() {
        var all = [];

        try {
            search.create({
                type: PROVIDER_LIST,
                filters: [
                    ['isinactive', 'is', 'F']
                ],
                columns: [
                    search.createColumn({
                        name: 'name'
                    })
                ]
            }).run().each(function (r) {
                all.push({
                    id: String(r.id),
                    name: String(r.getValue({
                        name: 'name'
                    }) || '')
                });

                return true;
            });

        } catch (e) {
            log.error('getProviders failed', e);
        }

        // Scope to the configured provider, if the deployment set one.
        // custscript_shipping_provider is a List/Record field, so it holds the
        // provider's internal id. Match by id.
        var configured = getConfiguredProvider();

        if (!configured) {
            return all; // empty param -> all providers
        }

        var wantedId = String(configured);

        var scoped = all.filter(function (p) {
            return String(p.id) === wantedId;
        });

        if (!scoped.length) {
            log.audit('Configured provider not found in active provider list', {
                param: SHIPPING_PROVIDER_PARAM,
                value: wantedId
            });
        }

        return scoped;
    }

    function getDefaultProviderId(providers) {
        return providers.length ? providers[0].id : '';
    }

    // Ensures a portal can only act on providers it is scoped to. If the
    // incoming provider id is not in this deployment's list, fall back to the
    // portal default. This keeps the TOSAN portal and the "others" portal from
    // operating on each other's data via a crafted request.
    function resolveProviderInScope(requestedId, providers) {
        requestedId = String(requestedId || '');

        if (requestedId) {
            for (var i = 0; i < providers.length; i++) {
                if (String(providers[i].id) === requestedId) {
                    return requestedId;
                }
            }
        }

        return getDefaultProviderId(providers);
    }

    function providerNameById(providers, id) {
        for (var i = 0; i < providers.length; i++) {
            if (String(providers[i].id) === String(id)) {
                return providers[i].name;
            }
        }

        return DEFAULT_NAME;
    }

    function getUserEmail() {
        try {
            return runtime.getCurrentUser().email || '';
        } catch (e) {
            return '';
        }
    }

    function getData(providerId, ctx) {
        ctx = ctx || {};

        if (!providerId) {
            return [];
        }

        try {
            return runSOSearch(providerId, true);

        } catch (e1) {
            log.audit('Zone column failed. Retrying without zone.', e1.message || e1);

            try {
                return runSOSearch(providerId, false);

            } catch (e2) {
                log.error('SO search failed', e2);
                ctx.error = 'Search error: ' + (e2.message || e2);
                return [];
            }
        }
    }

    function runSOSearch(providerId, includeZone) {

        // ------------------------------------------------------------------
        // MAIN SALES ORDER SEARCH (line-level, grouped per SO).
        //
        //   - mainline / taxline / shipping / cogs = F  (read the item lines)
        //   - Pallet  = SUM( NVL(item volume,0) * quantity )
        //   - Cases   = SUM( quantity )
        //   - Weight  = SUM( NVL(line weight,0) / 1000 )
        //
        // Header fields carried through (GROUP):
        //   - internalid  -> SO links + saving dates + POD upload target.
        //   - Expected Ship Date -> drives week / day allocation.
        //   - Delivery Zone -> drives zone grouping.
        //   - Shipping City / Postal Code -> display + export.
        //
        // POD files are NOT part of this grouped search (file attachments are
        // not a summary column). They are loaded separately in attachPodFiles().
        // ------------------------------------------------------------------

        var tranidCol = search.createColumn({
            name: 'tranid',
            summary: search.Summary.GROUP
        });

        var entityCol = search.createColumn({
            name: 'entity',
            summary: search.Summary.GROUP
        });

        var internalIdCol = search.createColumn({
            name: 'internalid',
            summary: search.Summary.GROUP
        });

        var shipDateCol = search.createColumn({
            name: DATE_FIELD,
            summary: search.Summary.GROUP
        });

      var availablePickupAfterCol = search.createColumn({
          name: AVAILABLE_PICKUP_AFTER_FIELD,
          summary: search.Summary.GROUP
     });

       var actualPickupDateCol = search.createColumn({
           name: ACTUAL_PICKUP_DATE_FIELD,
          summary: search.Summary.GROUP
      });

        var memoCol = search.createColumn({
            name: MEMO_FIELD,
            summary: search.Summary.GROUP
        });

        var deliveryNoteCol = search.createColumn({
            name: DELIVERY_NOTE_FIELD,
            summary: search.Summary.GROUP
        });

        var cityCol = search.createColumn({
            name: CITY_FIELD,
            summary: search.Summary.GROUP
        });

        var zipCol = search.createColumn({
            name: ZIP_FIELD,
            summary: search.Summary.GROUP
        });

        var palletCol = search.createColumn({
            name: 'formulanumeric',
            summary: search.Summary.SUM,
            formula: 'NVL({' + LINE_ITEM_VOLUME_FIELD + '},0)*{quantity}'
        });

        var casesCol = search.createColumn({
            name: 'formulanumeric',
            summary: search.Summary.SUM,
            formula: '{quantity}'
        });

        var weightCol = search.createColumn({
            name: 'formulanumeric',
            summary: search.Summary.SUM,
            formula: 'NVL({' + LINE_WEIGHT_FIELD + '},0)/1000'
        });

        var columns = [
            tranidCol,
            entityCol,
            internalIdCol,
            shipDateCol,
            availablePickupAfterCol,
            actualPickupDateCol,
            memoCol,
            deliveryNoteCol,
            cityCol,
            zipCol,
            palletCol,
            casesCol,
            weightCol
        ];

        var zoneCol = null;

        if (includeZone) {
            zoneCol = search.createColumn({
                name: ZONE_FIELD,
                summary: search.Summary.GROUP
            });

            columns.push(zoneCol);
        }

        var soSearch = search.create({
            type: search.Type.SALES_ORDER,
            settings: [{
                name: 'consolidationtype',
                value: 'ACCTTYPE'
            }],
            filters: [
                ['type', 'anyof', 'SalesOrd'],
                'AND',
                ['mainline', 'is', 'F'],
                'AND',
                ['taxline', 'is', 'F'],
                'AND',
                ['shipping', 'is', 'F'],
                'AND',
                ['cogs', 'is', 'F'],
                'AND',
                [PROVIDER_FIELD, 'anyof', providerId],
                'AND',
                [DATE_FIELD, 'onorafter', 'today']
            ],
            columns: columns
        });

        var rows = [];
        var paged = soSearch.runPaged({
            pageSize: 1000
        });

        paged.pageRanges.forEach(function (pr) {
            paged.fetch({
                index: pr.index
            }).data.forEach(function (r) {

                var rawDate = r.getValue(shipDateCol);
                var shipDate = parseNetSuiteDate(rawDate);
                var rawAvailablePickupAfter = r.getValue(availablePickupAfterCol);
                var availablePickupAfterDate = parseNetSuiteDate(rawAvailablePickupAfter);

                var rawActualPickupDate = r.getValue(actualPickupDateCol);
                var actualPickupDate = parseNetSuiteDate(rawActualPickupDate);

                if (!shipDate) {
                    return;
                }

                var dayNum = shipDate.getDay();

                // Skip weekend expected shipping dates
                if (dayNum === 0 || dayNum === 6) {
                    return;
                }

                var weekInfo = isoWeekInfo(shipDate);

                var zoneValue = zoneCol ? r.getValue(zoneCol) : '';
                var zoneText = zoneCol ? r.getText(zoneCol) : '';
                var zone = normalizeZone(zoneValue, zoneText);

                rows.push({
                    soId: String(r.getValue(internalIdCol) || ''),
                    tranid: String(r.getValue(tranidCol) || ''),
                    customer: String(r.getText(entityCol) || r.getValue(entityCol) || ''),
                    memo: String(r.getValue(memoCol) || ''),
                    deliveryNote: String(r.getValue(deliveryNoteCol) || ''),
                    city: String(r.getValue(cityCol) || ''),
                    zip: String(r.getValue(zipCol) || ''),
                    podFiles: [], // populated by attachPodFiles()
                    volume: numVal(r.getValue(palletCol)),
                    cases: numVal(r.getValue(casesCol)),
                    weight: numVal(r.getValue(weightCol)),
                    ymd: toYMD(shipDate),
                    availablePickupAfterYmd: availablePickupAfterDate ? toYMD(availablePickupAfterDate) : '',
                    actualPickupDateYmd: actualPickupDate ? toYMD(actualPickupDate) : '',
                    week: weekInfo.week,
                    weekYear: weekInfo.year,
                    weekKey: weekInfo.key,
                    weekLabel: 'Week ' + weekInfo.week,
                    zone: zone,
                    zoneLabel: getZoneLabel(zone),
                    dayNum: dayNum,
                    dayShort: DAY_SHORT[dayNum] || ''
                });
            });
        });

        attachPodFiles(rows);

        rows.sort(function (a, b) {
            return cmp(a.weekKey, b.weekKey) ||
                (zoneSortKey(a.zone) - zoneSortKey(b.zone)) ||
                cmp(a.ymd, b.ymd) ||
                cmp(a.tranid, b.tranid);
        });

        return rows;
    }

    // Loads the files attached to each SO (Communication > Files) and stores
    // them on row.podFiles. Done in batches because internalid 'anyof' has a
    // practical limit on how many ids can be passed in one filter.
    function attachPodFiles(rows) {
        if (!rows.length) {
            return;
        }

        var byId = {};
        var ids = [];

        rows.forEach(function (r) {
            r.podFiles = [];

            if (r.soId && !byId[r.soId]) {
                byId[r.soId] = r;
                ids.push(r.soId);
            }
        });

        if (!ids.length) {
            return;
        }

        var BATCH = 800;

        for (var i = 0; i < ids.length; i += BATCH) {
            var slice = ids.slice(i, i + BATCH);

            try {
                search.create({
                    type: search.Type.TRANSACTION,
                    filters: [
                        ['internalid', 'anyof', slice],
                        'AND',
                        ['mainline', 'is', 'T'],
                        'AND',
                        ['file.internalid', 'noneof', '@NONE@']
                    ],
                    columns: [
                        search.createColumn({ name: 'internalid' }),
                        search.createColumn({ name: 'internalid', join: 'file' }),
                        search.createColumn({ name: 'name', join: 'file' }),
                        search.createColumn({ name: 'url', join: 'file' })
                    ]
                }).run().each(function (res) {
                    var soId = String(res.getValue({ name: 'internalid' }) || '');
                    var row = byId[soId];

                    if (!row) {
                        return true;
                    }

                    var fid = String(res.getValue({ name: 'internalid', join: 'file' }) || '');

                    if (!fid) {
                        return true;
                    }

                    row.podFiles.push({
                        id: fid,
                        name: String(res.getValue({ name: 'name', join: 'file' }) || ''),
                        url: String(res.getValue({ name: 'url', join: 'file' }) || '')
                    });

                    return true;
                });

            } catch (e) {
                log.error('attachPodFiles failed (batch starting ' + i + ')', e);
            }
        }
    }

    function getResultValue(result, fieldId) {
        try {
            return result.getValue({
                name: fieldId
            });
        } catch (e1) {
            try {
                return result.getValue(fieldId);
            } catch (e2) {
                return '';
            }
        }
    }

    function getResultText(result, fieldId) {
        try {
            return result.getText({
                name: fieldId
            });
        } catch (e1) {
            try {
                return result.getText(fieldId);
            } catch (e2) {
                return '';
            }
        }
    }

    function getZoneLabel(zone) {
        zone = String(zone || '').trim();

        if (!zone || zone === '-') {
            return 'No Zone';
        }

        return ZONE_LABEL_MAP[zone] || ('Zone ' + zone);
    }

    function getZoneLabelClient(zone) {
        var map = {
            '2': 'Zone 2/GTA',
            '3': 'Zone 3/Ottawa',
            '4': 'Zone 4/Montreal',
            '5': 'Zone 5/Oshawa',
            '6': 'Zone 6/Barrie-NewMarket-Brantford',
            '7': 'Zone 7/Windsor',
            '8': 'Zone 8/Hamilton and Niagara Guelph-Georgetown',
            '9': 'Zone 9/London'
        };

        zone = String(zone || '').trim();

        if (!zone || zone === '-') {
            return 'No Zone';
        }

        return map[zone] || ('Zone ' + zone);
    }

    // build the option list used by the Zone dropdown (mirrors Week dropdown).
    function getZoneOptions() {
        var opts = [];

        ZONES.forEach(function (z) {
            var key = String(z);

            opts.push({
                zone: key,
                label: getZoneLabel(key)
            });
        });

        return opts;
    }

    /* ============================ ACTIONS ============================ */

  function doSave(payload) {
    if (!payload) {
        return 'No changes submitted.';
    }

    var edits;

    try {
        edits = JSON.parse(payload);
    } catch (e) {
        throw new Error('Invalid save payload.');
    }

    if (!edits || !edits.length) {
        return 'No changes submitted.';
    }

    var ok = 0;
    var fail = 0;
    var skippedWeekend = 0;

    edits.forEach(function (edit) {
        try {
            if (!edit || !edit.id) {
                return;
            }

            var values = {};
            var hasChanges = false;

            // Existing Expected Ship Date update
            if (typeof edit.shipdate !== 'undefined' && edit.shipdate !== null) {
                var newDate = dateFromYMD(edit.shipdate);

                if (!newDate) {
                    throw new Error('Invalid expected ship date: ' + edit.shipdate);
                }

                var day = newDate.getDay();

                if (day === 0 || day === 6) {
                    skippedWeekend++;
                    return;
                }

                values[DATE_FIELD] = newDate;
                hasChanges = true;
            }

            // New Available To Pick Up After update
            if (typeof edit.availablePickupAfter !== 'undefined') {
                if (edit.availablePickupAfter) {
                    var availablePickupAfterDate = dateFromYMD(edit.availablePickupAfter);

                    if (!availablePickupAfterDate) {
                        throw new Error('Invalid Available To Pick Up After date: ' + edit.availablePickupAfter);
                    }

                    values[AVAILABLE_PICKUP_AFTER_FIELD] = availablePickupAfterDate;
                } else {
                    values[AVAILABLE_PICKUP_AFTER_FIELD] = '';
                }

                hasChanges = true;
            }

            // New Actual Pick Up Date update
            if (typeof edit.actualPickupDate !== 'undefined') {
                if (edit.actualPickupDate) {
                    var actualPickupDate = dateFromYMD(edit.actualPickupDate);

                    if (!actualPickupDate) {
                        throw new Error('Invalid Actual Pick Up Date: ' + edit.actualPickupDate);
                    }

                    values[ACTUAL_PICKUP_DATE_FIELD] = actualPickupDate;
                } else {
                    values[ACTUAL_PICKUP_DATE_FIELD] = '';
                }

                hasChanges = true;
            }

            if (!hasChanges) {
                return;
            }

            record.submitFields({
                type: record.Type.SALES_ORDER,
                id: edit.id,
                values: values,
                options: {
                    enableSourcing: false,
                    ignoreMandatoryFields: true
                }
            });

            ok++;

        } catch (err) {
            fail++;
            log.error('Save failed for SO ' + (edit && edit.id), err);
        }
    });

    var msg = 'Saved ' + ok + ' update(s).';

    if (skippedWeekend) {
        msg += ' ' + skippedWeekend + ' weekend expected ship date(s) skipped.';
    }

    if (fail) {
        msg += ' ' + fail + ' failed. See script logs.';
    }

    return msg;
}

    // Clears the Expected Ship Date on a Sales Order. With the date removed the SO
    // no longer matches the "onorafter today" filter, so it drops off the schedule
    // on the next refresh.
    function doRemove(soId) {
        soId = String(soId || '').trim();

        if (!soId) {
            return 'No sales order specified.';
        }

        try {
            var values = {};
            values[DATE_FIELD] = ''; // empty string clears the date field

            record.submitFields({
                type: record.Type.SALES_ORDER,
                id: soId,
                values: values,
                options: {
                    enableSourcing: false,
                    ignoreMandatoryFields: true
                }
            });

            return 'Removed expected ship date from SO (id ' + soId + ').';

        } catch (e) {
            log.error('Remove failed for SO ' + soId, e);
            throw new Error('Could not remove expected ship date: ' + (e.message || e));
        }
    }

    // Creates a single POD file in the File Cabinet and attaches it to the SO,
    // so it appears under Communication > Files on the Sales Order record.
    // The client uploads ONE file per request to avoid POST size limits.
    function doUpload(req) {
        var soId = String(req.parameters.tosan_soid || '').trim();

        if (!soId) {
            return 'No sales order specified for upload.';
        }

        var fileName = String(req.parameters.tosan_filename || '').trim();
        var fileMime = String(req.parameters.tosan_filemime || '').trim();
        var fileData = String(req.parameters.tosan_filedata || ''); // base64, no data: prefix

        if (!fileName || !fileData) {
            return 'No file received.';
        }

        // The client sends raw base64. As a guard, strip any whitespace that a
        // form-decoder may have introduced (e.g. '+' decoded to ' ') and restore
        // base64-safe characters so the bytes are preserved exactly.
        fileData = sanitizeBase64(fileData);

        var folderId = getPodFolderId();
        var fileType = mapFileType(fileName, fileMime);

        try {
            var createOptions = {
                name: sanitizeName(fileName),
                fileType: fileType,
                folder: folderId
            };

            if (isTextFileType(fileType)) {
                // Text types (CSV, TXT) must be stored as decoded UTF-8 text.
                // Storing base64 against a text file type would save the base64
                // characters as the file content (gibberish). Decode first.
                createOptions.contents = encode.convert({
                    string: fileData,
                    inputEncoding: encode.Encoding.BASE_64,
                    outputEncoding: encode.Encoding.UTF_8
                });
                createOptions.encoding = file.Encoding.UTF_8;
            } else {
                // Binary types (images, PDF, Excel, Word, etc.): hand NetSuite the
                // base64 and the BASE_64 flag so it decodes back to the exact bytes.
                createOptions.contents = fileData;
                createOptions.encoding = file.Encoding.BASE_64;
            }

            var fileObj = file.create(createOptions);
            var newFileId = fileObj.save();

            // Attach file to the SO -> shows under Communication > Files.
            record.attach({
                record: { type: 'file', id: newFileId },
                to: { type: record.Type.SALES_ORDER, id: soId }
            });

            return 'Uploaded "' + sanitizeName(fileName) + '" to SO id ' + soId + '.';

        } catch (e) {
            log.error('Upload failed for SO ' + soId + ' file ' + fileName, e);
            throw new Error('Could not upload file: ' + (e.message || e));
        }
    }

    function getPodFolderId() {
        try {
            var p = runtime.getCurrentScript().getParameter({ name: POD_UPLOAD_FOLDER_PARAM });
            return p ? Number(p) : DEFAULT_POD_FOLDER_ID;
        } catch (e) {
            return DEFAULT_POD_FOLDER_ID;
        }
    }

    function sanitizeName(name) {
        return String(name || 'POD').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 200);
    }

    // Repairs base64 that may have been mangled by form-body decoding:
    // a '+' in the payload can be decoded to a space, so restore spaces to '+'.
    // Also strips any other whitespace (newlines/tabs) that should not be in base64.
    function sanitizeBase64(s) {
        s = String(s || '');
        s = s.replace(/ /g, '+');           // space -> '+' (form-decode artifact)
        s = s.replace(/[\r\n\t]/g, '');     // strip stray whitespace
        return s;
    }

    function isTextFileType(fileType) {
        return fileType === file.Type.CSV ||
            fileType === file.Type.PLAINTEXT ||
            fileType === file.Type.HTMLDOC ||
            fileType === file.Type.XMLDOC;
    }

    function mapFileType(name, mime) {
        var ext = String(name || '').split('.').pop().toLowerCase();

        switch (ext) {
            case 'pdf':  return file.Type.PDF;
            case 'jpg':
            case 'jpeg': return file.Type.JPGIMAGE;
            case 'png':  return file.Type.PNGIMAGE;
            case 'gif':  return file.Type.GIFIMAGE;
            case 'bmp':  return file.Type.BMPIMAGE;
            case 'tif':
            case 'tiff': return file.Type.TIFFIMAGE;
            case 'csv':  return file.Type.CSV;
            case 'xls':  return file.Type.EXCEL;
            case 'xlsx': return file.Type.EXCEL;
            case 'doc':  return file.Type.WORD;
            case 'docx': return file.Type.WORD;
            case 'txt':  return file.Type.PLAINTEXT;
            case 'heic': return file.Type.MISCBINARY;
            default:     return file.Type.MISCBINARY;
        }
    }

    function doEmail(providerId, providers, recipient) {
        recipient = String(recipient || '').trim();

        if (!recipient) {
            return 'No recipient email provided.';
        }

        var ctx = {};
        var data = getData(providerId, ctx);

        if (ctx.error) {
            throw new Error(ctx.error);
        }

        var providerName = providerNameById(providers, providerId);
        var html = buildEmailHtml(providerName, data);

        email.send({
            author: runtime.getCurrentUser().id,
            recipients: recipient.split(/[;,]/).map(function (s) {
                return s.trim();
            }).filter(Boolean),
            subject: 'TOSAN Shipping Schedule - ' + providerName,
            body: html
        });

        return 'Email sent to ' + recipient + '.';
    }

    function buildEmailHtml(providerName, data) {
        var grouped = groupRows(data);
        var t = totals(data);

        var html = '';

        html += '<div style="font-family:Arial,Helvetica,sans-serif;color:#111827">';
        html += '<h2 style="margin:0 0 6px">TOSAN - Weekly Shipping Schedule</h2>';
        html += '<p style="margin:0 0 14px">Provider: <b>' + esc(providerName) + '</b></p>';

        // per-zone summary (Total Orders / Pallets / Cases by Zone).
        html += '<h3 style="margin:0 0 6px">Totals by Zone</h3>';
        html += '<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;width:100%;font-size:12px;margin:0 0 18px">';
        html += '<thead>';
        html += '<tr style="background:#f3f4f6">';
        html += '<th align="left">Zone</th>';
        html += '<th align="right">Orders</th>';
        html += '<th align="right">Pallet</th>';
        html += '<th align="right">Cases</th>';
        html += '</tr>';
        html += '</thead>';
        html += '<tbody>';

        if (!t.byZone.length) {
            html += '<tr><td colspan="4" align="center">No sales orders found.</td></tr>';
        } else {
            t.byZone.forEach(function (z) {
                html += '<tr>';
                html += '<td>' + esc(z.zoneLabel) + '</td>';
                html += '<td align="right">' + fmt(z.orders) + '</td>';
                html += '<td align="right">' + fmt(z.volume) + '</td>';
                html += '<td align="right">' + fmt(z.cases) + '</td>';
                html += '</tr>';
            });

            html += '<tr style="background:#f3f4f6;font-weight:bold">';
            html += '<td>Grand Total</td>';
            html += '<td align="right">' + fmt(t.orders) + '</td>';
            html += '<td align="right">' + fmt(t.volume) + '</td>';
            html += '<td align="right">' + fmt(t.cases) + '</td>';
            html += '</tr>';
        }

        html += '</tbody>';
        html += '</table>';

        html += '<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;width:100%;font-size:12px">';
        html += '<thead>';
        html += '<tr style="background:#f3f4f6">';
        html += '<th align="left">Week</th>';
        html += '<th align="left">Zone</th>';
        html += '<th align="left">Day</th>';
        html += '<th align="left">Sales Order</th>';
        html += '<th align="left">Customer</th>';
        html += '<th align="left">City</th>';
        html += '<th align="left">Postal Code</th>';
        html += '<th align="right">Pallet</th>';
        html += '<th align="right">Cases</th>';
        html += '<th align="left">Expected Ship Date</th>';
        html += '<th align="left">Available To Pick Up After</th>';
        html += '<th align="left">Actual Pick Up Date</th>';
        html += '<th align="left">Memo</th>';
        html += '<th align="left">Delivery Note</th>';
        html += '<th align="left">POD Files</th>';
        html += '</tr>';
        html += '</thead>';
        html += '<tbody>';

        if (!data.length) {
            html += '<tr><td colspan="15" align="center">No sales orders found.</td></tr>';
        }

        Object.keys(grouped).sort().forEach(function (weekKey) {
            var week = grouped[weekKey];

            Object.keys(week.zones).sort(function (a, b) {
                return zoneSortKey(a) - zoneSortKey(b);
            }).forEach(function (zone) {
                week.zones[zone].forEach(function (r) {
                    html += '<tr>';
                    html += '<td>' + esc(r.weekLabel) + '</td>';
                    html += '<td>' + esc(r.zoneLabel) + '</td>';
                    html += '<td>' + esc(r.dayShort) + '</td>';
                    html += '<td>' + esc(r.tranid) + '</td>';
                    html += '<td>' + esc(r.customer) + '</td>';
                    html += '<td>' + esc(r.city) + '</td>';
                    html += '<td>' + esc(r.zip) + '</td>';
                    html += '<td align="right">' + fmt(r.volume) + '</td>';
                    html += '<td align="right">' + fmt(r.cases) + '</td>';
                    html += '<td>' + esc(r.ymd) + '</td>';
                    html += '<td>' + esc(r.availablePickupAfterYmd) + '</td>';
                    html += '<td>' + esc(r.actualPickupDateYmd) + '</td>';
                    html += '<td>' + esc(r.memo) + '</td>';
                    html += '<td>' + esc(r.deliveryNote) + '</td>';
                    html += '<td>' + buildPodLinksHtml(r.podFiles) + '</td>';
                    html += '</tr>';
                });
            });
        });

        html += '</tbody>';
        html += '</table>';
        html += '</div>';

        return html;
    }

    function buildPodLinksHtml(podFiles) {
        if (!podFiles || !podFiles.length) {
            return '';
        }

        return podFiles.map(function (f) {
            return '<a href="' + esc(f.url) + '">' + esc(f.name) + '</a>';
        }).join('<br>');
    }

    /* ============================ HELPERS ============================ */

    function getWeekWindow() {
        var today = cleanDate(new Date());

        // Current calendar year
        var start = new Date(today.getFullYear(), 0, 1);
        var end = new Date(today.getFullYear(), 11, 31);

        return {
            start: start,
            end: end,
            startYmd: toYMD(start),
            endYmd: toYMD(end),
            startStr: format.format({
                value: start,
                type: format.Type.DATE
            }),
            endStr: format.format({
                value: end,
                type: format.Type.DATE
            })
        };
    }

    function parseNetSuiteDate(value) {
        if (!value) {
            return null;
        }

        if (Object.prototype.toString.call(value) === '[object Date]') {
            return cleanDate(value);
        }

        var s = String(value).trim();

        if (!s) {
            return null;
        }

        try {
            return cleanDate(format.parse({
                value: s,
                type: format.Type.DATE
            }));
        } catch (e1) {
            // fallback below
        }

        var ymd = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

        if (ymd) {
            return cleanDate(new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3])));
        }

        var mdy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);

        if (mdy) {
            var y = Number(mdy[3]);

            if (y < 100) {
                y += 2000;
            }

            return cleanDate(new Date(y, Number(mdy[1]) - 1, Number(mdy[2])));
        }

        var d = new Date(s);

        if (!isNaN(d.getTime())) {
            return cleanDate(d);
        }

        return null;
    }

    function dateFromYMD(ymd) {
        var p = String(ymd || '').split('-');

        if (p.length !== 3) {
            return null;
        }

        var y = Number(p[0]);
        var m = Number(p[1]);
        var d = Number(p[2]);

        if (!y || !m || !d) {
            return null;
        }

        return cleanDate(new Date(y, m - 1, d));
    }

    function cleanDate(d) {
        return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }

    function isoWeekInfo(d) {
        var date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
        var dayNum = (date.getUTCDay() + 6) % 7;

        date.setUTCDate(date.getUTCDate() - dayNum + 3);

        var weekYear = date.getUTCFullYear();
        var firstThursday = new Date(Date.UTC(weekYear, 0, 4));
        var firstDayNum = (firstThursday.getUTCDay() + 6) % 7;

        firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNum + 3);

        var week = 1 + Math.round((date.getTime() - firstThursday.getTime()) / 604800000);
        var weekString = String(week);

        if (weekString.length < 2) {
            weekString = '0' + weekString;
        }

        return {
            week: week,
            year: weekYear,
            key: weekYear + '-W' + weekString,
            label: 'Week ' + week
        };
    }

    function normalizeZone(value, text) {
        var source = String(text || '').trim();

        if (!source) {
            source = String(value || '').trim();
        }

        if (!source) {
            return '-';
        }

        var match = source.match(/(?:zone\s*)?(\d+)/i);

        if (match) {
            return String(Number(match[1]));
        }

        return source;
    }

    function zoneSortKey(zone) {
        var n = parseInt(zone, 10);
        return isNaN(n) ? 999 : n;
    }

    function groupRows(data) {
        var grouped = {};

        data.forEach(function (r) {
            grouped[r.weekKey] = grouped[r.weekKey] || {
                rows: [],
                zones: {}
            };

            grouped[r.weekKey].rows.push(r);

            grouped[r.weekKey].zones[r.zone] = grouped[r.weekKey].zones[r.zone] || [];
            grouped[r.weekKey].zones[r.zone].push(r);
        });

        return grouped;
    }

    function totals(data) {
        var pallet = 0;
        var cases = 0;
        var weight = 0;
        var z = {};
        var byZoneMap = {};

        data.forEach(function (r) {
            var v = Number(r.volume) || 0;
            var c = Number(r.cases) || 0;
            var w = Number(r.weight) || 0;

            pallet += v;
            cases += c;
            weight += w;

            var zoneKey = (r.zone && r.zone !== '-') ? String(r.zone) : '-';

            if (zoneKey !== '-') {
                z[zoneKey] = true;
            }

            if (!byZoneMap[zoneKey]) {
                byZoneMap[zoneKey] = {
                    zone: zoneKey,
                    zoneLabel: getZoneLabel(zoneKey),
                    orders: 0,
                    volume: 0,
                    cases: 0,
                    weight: 0
                };
            }

            byZoneMap[zoneKey].orders += 1;
            byZoneMap[zoneKey].volume += v;
            byZoneMap[zoneKey].cases += c;
            byZoneMap[zoneKey].weight += w;
        });

        // per-zone totals (orders / pallets / cases / weight) sorted by zone number.
        var byZone = Object.keys(byZoneMap).sort(function (a, b) {
            return zoneSortKey(a) - zoneSortKey(b);
        }).map(function (k) {
            var e = byZoneMap[k];
            e.volume = round2(e.volume);
            e.cases = round2(e.cases);
            e.weight = round2(e.weight);
            return e;
        });

        return {
            orders: data.length,
            volume: round2(pallet),
            cases: round2(cases),
            weight: round2(weight),
            zones: Object.keys(z).length,
            byZone: byZone
        };
    }

    function numVal(value) {
        var n = parseFloat(String(value == null ? '' : value).replace(/,/g, ''));
        return isNaN(n) ? 0 : n;
    }

    function round2(n) {
        return Math.round((Number(n) || 0) * 100) / 100;
    }

    function fmt(n) {
        return round2(n).toLocaleString('en-US');
    }

    function toYMD(d) {
        return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
    }

    function pad2(n) {
        return ('0' + n).slice(-2);
    }

    function cmp(a, b) {
        a = String(a == null ? '' : a);
        b = String(b == null ? '' : b);

        if (a < b) {
            return -1;
        }

        if (a > b) {
            return 1;
        }

        return 0;
    }

    function esc(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function safeJsonForScript(obj) {
        return JSON.stringify(obj)
            .replace(/</g, '\\u003c')
            .replace(/>/g, '\\u003e')
            .replace(/&/g, '\\u0026')
            .replace(/\u2028/g, '\\u2028')
            .replace(/\u2029/g, '\\u2029');
    }

    function writeJson(res, payload) {
        res.setHeader({
            name: 'Content-Type',
            value: 'application/json; charset=utf-8'
        });

        res.write({
            output: JSON.stringify(payload || {})
        });
    }

    /* ==================== TRIP BUILDER: BACKEND ==================== */

    function getReserveTtlMin() {
        try {
            var p = runtime.getCurrentScript().getParameter({ name: RESERVE_TTL_PARAM });
            var n = Number(p);
            return (n && n > 0) ? n : DEFAULT_RESERVE_TTL_MIN;
        } catch (e) {
            return DEFAULT_RESERVE_TTL_MIN;
        }
    }

    // Unassigned orders for the trip builder: reuses the same SO search but adds
    // a "no trip" filter, then hides orders freshly reserved by OTHER users.
    function getUnassignedResponse(providerId, providers) {
        var rows = [];
        var errMsg = '';

        try {
            rows = runTripSOSearch(providerId, true);
        } catch (e1) {
            try {
                rows = runTripSOSearch(providerId, false);
            } catch (e2) {
                errMsg = 'Search error: ' + (e2.message || e2);
                rows = [];
            }
        }

        var meId = String(getCurrentUserId());
        var ttlMs = getReserveTtlMin() * 60 * 1000;
        var now = Date.now();

        var reserved = getReservations(rows.map(function (r) { return r.soId; }));

        var visible = rows.filter(function (r) {
            var info = reserved[r.soId];
            if (!info || !info.by) { return true; }
            var fresh = info.at && (now - info.at) < ttlMs;
            if (!fresh) { return true; }
            return String(info.by) === meId;
        });

        return {
            success: !errMsg,
            message: errMsg,
            selectedProvider: providerId,
            providerName: providerNameById(providers, providerId),
            zoneLabelMap: ZONE_LABEL_MAP,
            tempColors: TEMP_COLORS,
            reserveTtlMin: getReserveTtlMin(),
            trucks: getTrucks(),
            drivers: getDrivers(),
            myReserved: getMyReserved(),
            data: visible
        };
    }

    // A lean version of the SO search for the trip builder. Same line-grouped
    // logic; adds "no trip" filter and per-SO temp set. No POD load (not needed).
    function runTripSOSearch(providerId, includeZone) {
        var tranidCol = search.createColumn({ name: 'tranid', summary: search.Summary.GROUP });
        var entityCol = search.createColumn({ name: 'entity', summary: search.Summary.GROUP });
        var internalIdCol = search.createColumn({ name: 'internalid', summary: search.Summary.GROUP });
        var shipDateCol = search.createColumn({ name: DATE_FIELD, summary: search.Summary.GROUP });
        var cityCol = search.createColumn({ name: CITY_FIELD, summary: search.Summary.GROUP });
        var zipCol = search.createColumn({ name: ZIP_FIELD, summary: search.Summary.GROUP });

        var palletCol = search.createColumn({
            name: 'formulanumeric', summary: search.Summary.SUM,
            formula: 'NVL({' + LINE_ITEM_VOLUME_FIELD + '},0)*{quantity}'
        });
        var casesCol = search.createColumn({
            name: 'formulanumeric', summary: search.Summary.SUM, formula: '{quantity}'
        });
        var weightCol = search.createColumn({
            name: 'formulanumeric', summary: search.Summary.SUM,
            formula: 'NVL({' + LINE_WEIGHT_FIELD + '},0)/1000'
        });

        var columns = [tranidCol, entityCol, internalIdCol, shipDateCol, cityCol, zipCol, palletCol, casesCol, weightCol];

        var zoneCol = null;
        if (includeZone) {
            zoneCol = search.createColumn({ name: ZONE_FIELD, summary: search.Summary.GROUP });
            columns.push(zoneCol);
        }

        var soSearch = search.create({
            type: search.Type.SALES_ORDER,
            settings: [{ name: 'consolidationtype', value: 'ACCTTYPE' }],
            filters: [
                ['type', 'anyof', 'SalesOrd'],
                'AND', ['mainline', 'is', 'F'],
                'AND', ['taxline', 'is', 'F'],
                'AND', ['shipping', 'is', 'F'],
                'AND', ['cogs', 'is', 'F'],
                'AND', [PROVIDER_FIELD, 'anyof', providerId],
                'AND', [DATE_FIELD, 'onorafter', 'today'],
                'AND', [SO_TRIP_FIELD, 'anyof', '@NONE@']
            ],
            columns: columns
        });

        var rows = [];
        var paged = soSearch.runPaged({ pageSize: 1000 });

        paged.pageRanges.forEach(function (pr) {
            paged.fetch({ index: pr.index }).data.forEach(function (r) {
                var shipDate = parseNetSuiteDate(r.getValue(shipDateCol));
                if (!shipDate) { return; }
                var dayNum = shipDate.getDay();
                if (dayNum === 0 || dayNum === 6) { return; }

                var weekInfo = isoWeekInfo(shipDate);
                var zoneValue = zoneCol ? r.getValue(zoneCol) : '';
                var zoneText = zoneCol ? r.getText(zoneCol) : '';
                var zone = normalizeZone(zoneValue, zoneText);

                rows.push({
                    soId: String(r.getValue(internalIdCol) || ''),
                    tranid: String(r.getValue(tranidCol) || ''),
                    customer: String(r.getText(entityCol) || r.getValue(entityCol) || ''),
                    city: String(r.getValue(cityCol) || ''),
                    zip: String(r.getValue(zipCol) || ''),
                    volume: numVal(r.getValue(palletCol)),
                    cases: numVal(r.getValue(casesCol)),
                    weight: numVal(r.getValue(weightCol)),
                    tempSet: [],
                    ymd: toYMD(shipDate),
                    week: weekInfo.week,
                    weekKey: weekInfo.key,
                    weekLabel: 'Week ' + weekInfo.week,
                    zone: zone,
                    zoneLabel: getZoneLabel(zone),
                    dayNum: dayNum,
                    dayShort: DAY_SHORT[dayNum] || ''
                });
            });
        });

        attachTempSets(rows);

        rows.sort(function (a, b) {
            return cmp(a.weekKey, b.weekKey) ||
                (zoneSortKey(a.zone) - zoneSortKey(b.zone)) ||
                cmp(a.ymd, b.ymd) || cmp(a.tranid, b.tranid);
        });

        return rows;
    }

    function attachTempSets(rows) {
        if (!rows.length) { return; }

        var ids = [];
        rows.forEach(function (r) { r.tempSet = []; if (r.soId) { ids.push(r.soId); } });
        if (!ids.length) { return; }

        var setMap = {};
        var BATCH = 800;

        for (var i = 0; i < ids.length; i += BATCH) {
            var slice = ids.slice(i, i + BATCH);
            try {
                // custitem8 lives on the Item record; reference it with a proper
                // join column ({ name:'custitem8', join:'item' }) rather than the
                // dotted string 'item.custitem8' which NetSuite rejects as invalid.
                var idCol = search.createColumn({ name: 'internalid', summary: search.Summary.GROUP });
                var tempCol = search.createColumn({ name: 'custitem8', join: 'item', summary: search.Summary.GROUP });

                search.create({
                    type: search.Type.SALES_ORDER,
                    filters: [
                        ['internalid', 'anyof', slice],
                        'AND', ['mainline', 'is', 'F'],
                        'AND', ['taxline', 'is', 'F'],
                        'AND', ['shipping', 'is', 'F'],
                        'AND', ['cogs', 'is', 'F']
                    ],
                    columns: [idCol, tempCol]
                }).run().each(function (res) {
                    var soId = String(res.getValue(idCol) || '');
                    var key = normalizeTemp(res.getValue(tempCol));
                    if (soId && key) { setMap[soId] = setMap[soId] || {}; setMap[soId][key] = true; }
                    return true;
                });
            } catch (e) {
                log.error('attachTempSets failed (batch ' + i + ')', e);
            }
        }

        rows.forEach(function (r) {
            var keys = setMap[r.soId] ? orderTempKeys(Object.keys(setMap[r.soId])) : [];
            r.tempSet = keys;
        });
    }

    function normalizeTemp(raw) {
        var s = String(raw == null ? '' : raw).trim().toUpperCase();
        if (!s) { return ''; }
        if (s.indexOf('TC/R') !== -1) { return TEMP_TCR; }
        if (s.indexOf('TC') !== -1) { return TEMP_TC; }
        if (s.charAt(0) === 'R') { return TEMP_R; }
        if (s.indexOf('AMBIENT') !== -1) { return TEMP_AMBIENT; }
        return TEMP_AMBIENT;
    }

    function orderTempKeys(keys) {
        var order = [TEMP_AMBIENT, TEMP_R, TEMP_TC, TEMP_TCR];
        return keys.slice().sort(function (a, b) { return order.indexOf(a) - order.indexOf(b); });
    }

    function tempSetHasConflict(keys) {
        var hasR = keys.indexOf(TEMP_R) !== -1;
        var hasTC = keys.indexOf(TEMP_TC) !== -1;
        var hasTCR = keys.indexOf(TEMP_TCR) !== -1;
        return hasR && hasTC && !hasTCR;
    }

    function getReservations(soIds) {
        var out = {};
        var ids = (soIds || []).filter(Boolean);
        if (!ids.length) { return out; }

        var BATCH = 800;
        for (var i = 0; i < ids.length; i += BATCH) {
            var slice = ids.slice(i, i + BATCH);
            try {
                search.create({
                    type: search.Type.SALES_ORDER,
                    filters: [
                        ['internalid', 'anyof', slice],
                        'AND', ['mainline', 'is', 'T'],
                        'AND', [SO_RESERVED_BY_FIELD, 'noneof', '@NONE@']
                    ],
                    columns: [
                        search.createColumn({ name: 'internalid' }),
                        search.createColumn({ name: SO_RESERVED_BY_FIELD }),
                        search.createColumn({ name: SO_RESERVED_AT_FIELD })
                    ]
                }).run().each(function (res) {
                    var soId = String(res.getValue({ name: 'internalid' }) || '');
                    var by = String(res.getValue({ name: SO_RESERVED_BY_FIELD }) || '');
                    var atDate = parseNetSuiteDateTime(res.getValue({ name: SO_RESERVED_AT_FIELD }));
                    out[soId] = { by: by, at: atDate ? atDate.getTime() : 0 };
                    return true;
                });
            } catch (e) {
                log.error('getReservations failed (batch ' + i + ')', e);
            }
        }
        return out;
    }

    function getCurrentUserId() {
        try { return runtime.getCurrentUser().id || ''; } catch (e) { return ''; }
    }

    function getTrucks() {
        var list = [];
        try {
            search.create({
                type: TRUCK_LIST,
                filters: [['isinactive', 'is', 'F']],
                columns: [search.createColumn({ name: 'name' })]
            }).run().each(function (r) {
                list.push({
                    id: String(r.id),
                    name: String(r.getValue({ name: 'name' }) || ''),
                    // A plain custom list has no capability data.
                    tempControl: false,
                    multitemp: false
                });
                return true;
            });
        } catch (e) {
            log.error('getTrucks failed', e);
        }
        list.sort(function (a, b) { return cmp(a.name, b.name); });
        return list;
    }

    // Returns the SOs currently reserved by the CURRENT user (fresh reservations
    // only, within TTL). Used by the "My reserved orders" panel so a planner can
    // see and release what they are holding across sessions.
    function getMyReserved() {
        var meId = String(getCurrentUserId());
        if (!meId) { return []; }

        var ttlMs = getReserveTtlMin() * 60 * 1000;
        var now = Date.now();
        var out = [];

        try {
            search.create({
                type: search.Type.SALES_ORDER,
                filters: [
                    ['mainline', 'is', 'T'],
                    'AND', [SO_RESERVED_BY_FIELD, 'anyof', meId],
                    'AND', [SO_TRIP_FIELD, 'anyof', '@NONE@']
                ],
                columns: [
                    search.createColumn({ name: 'internalid' }),
                    search.createColumn({ name: 'tranid' }),
                    search.createColumn({ name: 'entity' }),
                    search.createColumn({ name: SO_RESERVED_AT_FIELD })
                ]
            }).run().each(function (r) {
                var atDate = parseNetSuiteDateTime(r.getValue({ name: SO_RESERVED_AT_FIELD }));
                var atMs = atDate ? atDate.getTime() : 0;
                // Only show fresh reservations; stale ones auto-release anyway.
                if (atMs && (now - atMs) >= ttlMs) { return true; }

                out.push({
                    soId: String(r.getValue({ name: 'internalid' }) || ''),
                    tranid: String(r.getValue({ name: 'tranid' }) || ''),
                    customer: String(r.getText({ name: 'entity' }) || r.getValue({ name: 'entity' }) || ''),
                    reservedAtMs: atMs
                });
                return true;
            });
        } catch (e) {
            log.error('getMyReserved failed', e);
        }

        return out;
    }

    // Loads the driver custom list (id + name) for the trip Driver dropdown.
    function getDrivers() {
        var list = [];
        try {
            search.create({
                type: DRIVER_LIST,
                filters: [['isinactive', 'is', 'F']],
                columns: [search.createColumn({ name: 'name' })]
            }).run().each(function (r) {
                list.push({
                    id: String(r.id),
                    name: String(r.getValue({ name: 'name' }) || '')
                });
                return true;
            });
        } catch (e) {
            log.error('getDrivers failed', e);
        }
        list.sort(function (a, b) { return cmp(a.name, b.name); });
        return list;
    }

    function reserveSo(soId) {
        soId = String(soId || '').trim();
        if (!soId) { return { success: false, message: 'No sales order specified.' }; }

        var check = getSoTripAndReservation(soId);
        if (check.trip) { return { success: false, message: 'Order is already on a trip.' }; }

        var meId = String(getCurrentUserId());
        var ttlMs = getReserveTtlMin() * 60 * 1000;
        var now = Date.now();

        if (check.reservedBy && String(check.reservedBy) !== meId) {
            if (check.reservedAt && (now - check.reservedAt) < ttlMs) {
                return { success: false, message: 'Order is being planned by another user.' };
            }
        }

        try {
            var values = {};
            values[SO_RESERVED_BY_FIELD] = meId;
            values[SO_RESERVED_AT_FIELD] = new Date();

            log.debug('reserveSo attempt', {
                soId: soId, meId: meId,
                byField: SO_RESERVED_BY_FIELD, atField: SO_RESERVED_AT_FIELD
            });

            record.submitFields({
                type: record.Type.SALES_ORDER, id: soId, values: values,
                options: { enableSourcing: false, ignoreMandatoryFields: true }
            });

            // Read back to confirm the write persisted. If the field ids are wrong
            // or permissions block the write, this catches it instead of silently
            // showing the card as reserved in the UI while nothing saved.
            var confirmRes = getSoTripAndReservation(soId);
            if (!confirmRes.reservedBy) {
                log.error('reserveSo wrote but read-back shows NO reservation', { soId: soId });
                return { success: false, message: 'Reserve did not persist on SO ' + soId + '. Check reserve field ids / permissions.' };
            }

            return { success: true, soId: soId };
        } catch (e) {
            log.error('reserveSo failed for SO ' + soId, e);
            return { success: false, message: 'Could not reserve SO ' + soId + ': ' + (e.message || e) };
        }
    }

    function releaseSo(soId) {
        soId = String(soId || '').trim();
        if (!soId) { return { success: false, message: 'No sales order specified.' }; }

        var check = getSoTripAndReservation(soId);
        var meId = String(getCurrentUserId());
        if (check.reservedBy && String(check.reservedBy) !== meId) {
            return { success: true, soId: soId };
        }

        try {
            var values = {};
            values[SO_RESERVED_BY_FIELD] = '';
            values[SO_RESERVED_AT_FIELD] = '';
            record.submitFields({
                type: record.Type.SALES_ORDER, id: soId, values: values,
                options: { enableSourcing: false, ignoreMandatoryFields: true }
            });
            return { success: true, soId: soId };
        } catch (e) {
            log.error('releaseSo failed for SO ' + soId, e);
            return { success: false, message: 'Could not release: ' + (e.message || e) };
        }
    }

    function getSoTripAndReservation(soId) {
        var out = { trip: '', reservedBy: '', reservedAt: 0 };
        try {
            var lk = search.lookupFields({
                type: search.Type.SALES_ORDER, id: soId,
                columns: [SO_TRIP_FIELD, SO_RESERVED_BY_FIELD, SO_RESERVED_AT_FIELD]
            });
            var tripVal = lk[SO_TRIP_FIELD];
            if (tripVal && tripVal.length) { out.trip = String(tripVal[0].value || ''); }
            else if (tripVal) { out.trip = String(tripVal || ''); }

            var byVal = lk[SO_RESERVED_BY_FIELD];
            if (byVal && byVal.length) { out.reservedBy = String(byVal[0].value || ''); }
            else if (byVal) { out.reservedBy = String(byVal || ''); }

            var atDate = parseNetSuiteDateTime(lk[SO_RESERVED_AT_FIELD]);
            out.reservedAt = atDate ? atDate.getTime() : 0;
        } catch (e) {
            log.error('getSoTripAndReservation failed for SO ' + soId, e);
        }
        return out;
    }

    function submitTrip(req, providerId) {
        var payload = req.parameters.tosan_payload;
        if (!payload) { return { success: false, message: 'No trip data submitted.' }; }

        var trip;
        try { trip = JSON.parse(payload); }
        catch (e) { return { success: false, message: 'Invalid trip payload.' }; }

        var orders = (trip && trip.orders) || [];
        if (!orders.length) { return { success: false, message: 'Add at least one order to the trip.' }; }

        var meId = String(getCurrentUserId());
        var ttlMs = getReserveTtlMin() * 60 * 1000;
        var now = Date.now();

        var usable = [];
        var dropped = [];

        orders.forEach(function (o) {
            var soId = String(o.soId || '').trim();
            if (!soId) { return; }
            var check = getSoTripAndReservation(soId);
            if (check.trip) { dropped.push({ soId: soId, reason: 'already on a trip' }); return; }
            if (check.reservedBy && String(check.reservedBy) !== meId) {
                if (check.reservedAt && (now - check.reservedAt) < ttlMs) {
                    dropped.push({ soId: soId, reason: 'reserved by another user' }); return;
                }
            }
            usable.push(o);
        });

        if (!usable.length) {
            return { success: false, message: 'None of the orders could be assigned.' };
        }

        var roll = rollupTrip(usable);

        try {
            var tripRec = record.create({ type: TRIP_TRANTYPE, isDynamic: true });


            tripRec.setValue({ fieldId: 'entity', value: TRIP_CUSTOMER_ID });
            tripRec.setValue({ fieldId: 'location', value: TRIP_LOCATION_ID });

            tripTrySetValue(tripRec, TRIP_PROVIDER_FIELD, providerId);
            if (trip.truckId) { tripTrySetValue(tripRec, TRIP_TRUCK_FIELD, trip.truckId); }
            if (trip.driver) { tripTrySetValue(tripRec, TRIP_DRIVER_FIELD, String(trip.driver)); }
            if (trip.shipDate) {
                var sd = dateFromYMD(trip.shipDate);
                if (sd) { tripTrySetValue(tripRec, TRIP_SHIP_DATE_FIELD, sd); }
            }

            tripTrySetValue(tripRec, TRIP_OVERRIDE_TOTALS_FIELD, false);
            tripTrySetValue(tripRec, TRIP_TOTAL_STOPS_FIELD, roll.stops);
            tripTrySetValue(tripRec, TRIP_TOTAL_PALLETS_FIELD, roll.pallets);
            tripTrySetValue(tripRec, TRIP_TOTAL_CASES_FIELD, roll.cases);
            tripTrySetValue(tripRec, TRIP_TOTAL_WEIGHT_FIELD, roll.weight);
            tripTrySetValue(tripRec, TRIP_TEMP_CONTROL_FIELD, roll.tempControl);
            tripTrySetValue(tripRec, TRIP_TEMP_SET_FIELD, roll.tempKeys.join(','));
            tripTrySetValue(tripRec, TRIP_TEMP_CONFLICT_FIELD, roll.tempConflict);

            // ITEMIZED MANIFEST: one trip line per SO item line.
            // Pull all item lines for the trip's SOs in one batched search, then
            // create a trip line per item carrying item + qty + rate, tagged with
            // the owning SO in custcol_trip_so. Non-item lines are skipped.
            var soIds = usable.map(function (o) { return o.soId; });
            var itemLinesBySo = getSoItemLines(soIds);

            // Diagnostic: how many item lines came back per SO.
            var lineCounts = {};
            Object.keys(itemLinesBySo).forEach(function (k) { lineCounts[k] = itemLinesBySo[k].length; });
            log.debug('getSoItemLines result', { soIds: soIds, counts: lineCounts });

            // Map for quick access to the trip-level snapshot per SO (zone/temp/stopSeq).
            var soMeta = {};
            usable.forEach(function (o, idx) {
                soMeta[String(o.soId)] = {
                    stopSeq: numVal(o.stopSeq || (idx + 1)),
                    zone: String(o.zone || ''),
                    pallets: numVal(o.pallets),
                    cases: numVal(o.cases),
                    weight: numVal(o.weight),
                    tempKeys: orderTempKeys(o.tempSet || [])
                };
            });

            var anyLineWritten = false;
            var droppedItems = [];

            usable.forEach(function (o) {
                var meta = soMeta[String(o.soId)] || {};
                var lines = itemLinesBySo[String(o.soId)] || [];

                lines.forEach(function (li) {
                    if (!li.item) { return; }

                    tripRec.selectNewLine({ sublistId: TRIP_LINE_SUBLIST });

                    // Set the item with a HARD call so a rejection surfaces. If the
                    // item is invalid for this transaction (inactive, wrong type,
                    // not sold, etc.), skip the line cleanly and keep going so the
                    // trip still saves with the valid items.
                    var itemSet = false;
                    try {
                        tripRec.setCurrentSublistValue({
                            sublistId: TRIP_LINE_SUBLIST,
                            fieldId: 'item',
                            value: li.item
                        });
                        itemSet = true;
                    } catch (itemErr) {
                        log.error('Set item failed on trip line (skipping)', {
                            sublist: TRIP_LINE_SUBLIST, item: li.item, soId: o.soId,
                            error: (itemErr && itemErr.message) || itemErr
                        });
                    }

                    if (!itemSet) {
                        droppedItems.push({ soId: String(o.soId), item: String(li.item) });
                        // Abandon this uncommitted line so it can't corrupt the next.
                        return;
                    }

                    tripTrySetSublist(tripRec, 'quantity', 0);
                    tripTrySetSublist(tripRec, TRIP_LINE_CUSTOM_QTY_FIELD, numVal(li.quantity));
                    tripTrySetSublist(tripRec, 'rate', numVal(li.rate));
                    tripTrySetSublist(tripRec, 'amount', numVal(li.amount));
                    // Line location comes from the SO line; fall back to the header
                    // location if the SO line had none.
                    tripTrySetSublist(tripRec, 'location', li.location || TRIP_LOCATION_ID);

                    // Tag which SO this item belongs to + carry trip snapshot fields.
                    tripTrySetSublist(tripRec, TRIP_LINE_SO_FIELD, String(o.soId));
                    tripTrySetSublist(tripRec, TRIP_LINE_STOP_SEQ_FIELD, meta.stopSeq);
                    tripTrySetSublist(tripRec, TRIP_LINE_ZONE_FIELD, meta.zone);
                    tripTrySetSublist(tripRec, TRIP_LINE_PALLETS_FIELD, meta.pallets);
                    tripTrySetSublist(tripRec, TRIP_LINE_CASES_FIELD, meta.cases);
                    tripTrySetSublist(tripRec, TRIP_LINE_WEIGHT_FIELD, meta.weight);
                    tripTrySetSublist(tripRec, TRIP_LINE_TEMP_SET_FIELD, (meta.tempKeys || []).join(','));
                    tripTrySetSublist(tripRec, TRIP_LINE_TEMP_CONTROL_FIELD,
                        (meta.tempKeys || []).some(function (k) { return k !== TEMP_AMBIENT; }));

                    try {
                        tripRec.commitLine({ sublistId: TRIP_LINE_SUBLIST });
                        anyLineWritten = true;
                    } catch (commitErr) {
                        log.error('commitLine failed on trip line', {
                            sublist: TRIP_LINE_SUBLIST, item: li.item, soId: o.soId,
                            error: (commitErr && commitErr.message) || commitErr
                        });
                    }
                });
            });

            // If no item lines were written at all (none of the SOs had readable
            // items), fail clearly rather than saving an empty trip.
            if (!anyLineWritten) {
                return { success: false, message: 'No item lines found on the selected orders. Trip not created.' };
            }

            var tripId = tripRec.save({ enableSourcing: true, ignoreMandatoryFields: true });

            var stamped = 0;
            var stampedLines = 0;
            usable.forEach(function (o) {
                try {
                    var stampResult = stampTripOnSalesOrder(o.soId, tripId);
                    stampedLines += stampResult.lines;
                    stamped++;
                } catch (e) {
                    log.error('Failed to stamp trip on SO ' + o.soId, e);
                }
            });

            var msg = 'Trip created (id ' + tripId + ') with ' + stamped + ' order(s).';
            if (stampedLines) { msg += ' ' + stampedLines + ' sales order line(s) linked.'; }
            if (dropped.length) { msg += ' ' + dropped.length + ' skipped.'; }
            if (droppedItems.length) {
                msg += ' ' + droppedItems.length + ' item line(s) could not be added (invalid for this transaction) — see script logs.';
            }
            if (roll.tempConflict) { msg += ' NOTE: temperature conflict \u2014 verify the truck.'; }

            return { success: true, message: msg, tripId: tripId, dropped: dropped };

        } catch (e) {
            log.error('submitTrip failed', e);
            return { success: false, message: 'Could not create trip: ' + (e.message || e) };
        }
    }

    function stampTripOnSalesOrder(soId, tripId) {
        var soRec = record.load({
            type: record.Type.SALES_ORDER,
            id: soId,
            isDynamic: false
        });

        soRec.setValue({ fieldId: SO_TRIP_FIELD, value: tripId });
        soRec.setValue({ fieldId: SO_RESERVED_BY_FIELD, value: '' });
        soRec.setValue({ fieldId: SO_RESERVED_AT_FIELD, value: '' });

        var lineCount = soRec.getLineCount({ sublistId: 'item' }) || 0;
        var lines = 0;

        for (var i = 0; i < lineCount; i++) {
            var itemId = soRec.getSublistValue({
                sublistId: 'item',
                fieldId: 'item',
                line: i
            });

            if (!itemId) {
                continue;
            }

            soRec.setSublistValue({
                sublistId: 'item',
                fieldId: SO_LINE_RELATED_TRIP_FIELD,
                line: i,
                value: tripId
            });

            lines++;
        }

        soRec.save({
            enableSourcing: false,
            ignoreMandatoryFields: true
        });

        return { lines: lines };
    }

    // Returns { soId: [ { item, quantity, rate }, ... ] } for the given SOs.
    // One batched line-level search; skips shipping / discount / tax / subtotal
    // and any line without an item.
    function getSoItemLines(soIds) {
        var out = {};
        var ids = (soIds || []).filter(Boolean);
        if (!ids.length) { return out; }

        var BATCH = 400;

        for (var i = 0; i < ids.length; i += BATCH) {
            var slice = ids.slice(i, i + BATCH);

            try {
                search.create({
                    type: search.Type.SALES_ORDER,
                    filters: [
                        ['internalid', 'anyof', slice],
                        'AND', ['mainline', 'is', 'F'],
                        'AND', ['taxline', 'is', 'F'],
                        'AND', ['shipping', 'is', 'F'],
                        'AND', ['cogs', 'is', 'F'],
                        // item must be a real item (excludes discount/subtotal rows)
                        'AND', ['item.type', 'noneof', 'Discount', 'Subtotal', 'Description', 'Payment']
                    ],
                    columns: [
                        search.createColumn({ name: 'internalid' }),
                        search.createColumn({ name: 'item' }),
                        search.createColumn({ name: 'quantity' }),
                        search.createColumn({ name: 'rate' }),
                        search.createColumn({ name: 'amount' }),
                        search.createColumn({ name: 'location' }),
                        search.createColumn({ name: 'line' })
                    ]
                }).run().each(function (r) {
                    var soId = String(r.getValue({ name: 'internalid' }) || '');
                    var itemId = String(r.getValue({ name: 'item' }) || '');
                    if (!soId || !itemId) { return true; }

                    out[soId] = out[soId] || [];
                    out[soId].push({
                        item: itemId,
                        quantity: numVal(r.getValue({ name: 'quantity' })),
                        rate: numVal(r.getValue({ name: 'rate' })),
                        amount: numVal(r.getValue({ name: 'amount' })),
                        location: String(r.getValue({ name: 'location' }) || '')
                    });
                    return true;
                });
            } catch (e) {
                log.error('getSoItemLines failed (batch ' + i + ')', e);
                // Fallback: retry this batch without the item.type filter, in case
                // that join is not supported; better to include a few extra rows
                // than to lose all item lines for the SO.
                try {
                    search.create({
                        type: search.Type.SALES_ORDER,
                        filters: [
                            ['internalid', 'anyof', slice],
                            'AND', ['mainline', 'is', 'F'],
                            'AND', ['taxline', 'is', 'F'],
                            'AND', ['shipping', 'is', 'F'],
                            'AND', ['cogs', 'is', 'F']
                        ],
                        columns: [
                            search.createColumn({ name: 'internalid' }),
                            search.createColumn({ name: 'item' }),
                            search.createColumn({ name: 'quantity' }),
                            search.createColumn({ name: 'rate' }),
                            search.createColumn({ name: 'location' })
                        ]
                    }).run().each(function (r) {
                        var soId = String(r.getValue({ name: 'internalid' }) || '');
                        var itemId = String(r.getValue({ name: 'item' }) || '');
                        if (!soId || !itemId) { return true; }
                        out[soId] = out[soId] || [];
                        out[soId].push({
                            item: itemId,
                            quantity: numVal(r.getValue({ name: 'quantity' })),
                            rate: numVal(r.getValue({ name: 'rate' })),
                            location: String(r.getValue({ name: 'location' }) || '')
                        });
                        return true;
                    });
                } catch (e2) {
                    log.error('getSoItemLines fallback failed (batch ' + i + ')', e2);
                }
            }
        }

        return out;
    }

    function rollupTrip(orders) {
        var pallets = 0, cases = 0, weight = 0;
        var keyMap = {};
        orders.forEach(function (o) {
            pallets += numVal(o.pallets);
            cases += numVal(o.cases);
            weight += numVal(o.weight);
            (o.tempSet || []).forEach(function (k) {
                var norm = normalizeTemp(k);
                if (norm) { keyMap[norm] = true; }
            });
        });
        var keys = orderTempKeys(Object.keys(keyMap));
        return {
            stops: orders.length,
            pallets: round2(pallets),
            cases: round2(cases),
            weight: round2(weight),
            tempKeys: keys,
            tempControl: keys.some(function (k) { return k !== TEMP_AMBIENT; }),
            tempConflict: tempSetHasConflict(keys)
        };
    }

    function tripTrySetValue(rec, fieldId, value) {
        try { rec.setValue({ fieldId: fieldId, value: value }); }
        catch (e) { log.error('trip setValue failed for ' + fieldId, e); }
    }

    function tripTrySetSublist(rec, fieldId, value) {
        try { rec.setCurrentSublistValue({ sublistId: TRIP_LINE_SUBLIST, fieldId: fieldId, value: value }); }
        catch (e) { log.error('trip setCurrentSublistValue failed for ' + fieldId, e); }
    }

    function parseNetSuiteDateTime(value) {
        if (!value) { return null; }
        if (Object.prototype.toString.call(value) === '[object Date]') { return value; }
        var s = String(value).trim();
        if (!s) { return null; }
        try { return format.parse({ value: s, type: format.Type.DATETIME }); }
        catch (e1) { /* fall through */ }
        var d = new Date(s);
        return isNaN(d.getTime()) ? null : d;
    }

    return {
        onRequest: onRequest
    };
});
