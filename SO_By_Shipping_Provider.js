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
 * - Filter / Save / Email uses AJAX, no full page refresh.
 *
 * Required Script Parameter:
 * custscript_tosan_html_file_id = File Cabinet internal ID or path of tosan_shipping_provider.html
 */
define([
    'N/search',
    'N/record',
    'N/email',
    'N/runtime',
    'N/format',
    'N/file',
    'N/log',
    'N/url'
], function (
    search,
    record,
    email,
    runtime,
    format,
    file,
    log,
    url
) {

    /* ============================= CONFIG ============================= */

    var PROVIDER_LIST = 'customlist_mi_shipping_provider_list';
    var PROVIDER_FIELD = 'custbody_shipping_provider_transaction';
    var ZONE_FIELD = 'custbody_mi_delivery_zone';
    var DATE_FIELD = 'custbody_expected_shipping_date';

    // UI Label = Pallet
    var VOLUME_FIELD = 'custbody_total_so_volume';

    // UI Label = Cases
    var CASES_FIELD = 'custbody_total_cases_to_ship';

    var DEFAULT_NAME = 'TOSAN SHIPPING PROVIDER';

    var HTML_FILE_PARAM = 'custscript_tosan_html_file_id';
    var DEFAULT_HTML_PATH = 'SuiteScripts/tosan_shipping_provider.html';

    var ZONES = [2, 3, 4, 5, 6, 7, 8, 9, 10];

    var DAY_LABELS = {
        1: 'Mon-1',
        2: 'Tue-2',
        3: 'Wed-3',
        4: 'Thu-4',
        5: 'Fri-5'
    };

    var DAY_SHORT = {
        1: 'MON',
        2: 'TUE',
        3: 'WED',
        4: 'THU',
        5: 'FRI'
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
            var selectedProvider = req.parameters.tosan_provider || getDefaultProviderId(providers);
            var recipient = req.parameters.tosan_recipient || getUserEmail();

            var result;

            if (action === 'data') {
                result = getDataResponse(selectedProvider, providers);
            } else if (action === 'save') {
                result = saveAndRefresh(req.parameters.tosan_payload, selectedProvider, providers);
            } else if (action === 'email') {
                result = emailAndRefresh(selectedProvider, providers, recipient);
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
            dayLabels: DAY_LABELS,
            dayShort: DAY_SHORT,
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

        var bootstrap = {
            apiUrl: getSuiteletUrl(),
            providers: providers,
            selectedProvider: selectedProvider,
            selectedProviderName: providerNameById(providers, selectedProvider),
            recipient: getUserEmail(),
            currentWeek: isoWeekInfo(new Date()),
            zones: ZONES,
            dayLabels: DAY_LABELS,
            dayShort: DAY_SHORT
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

    function getSuiteletUrl() {
        try {
            var scriptObj = runtime.getCurrentScript();

            return url.resolveScript({
                scriptId: scriptObj.id,
                deploymentId: scriptObj.deploymentId
            });

        } catch (e) {
            log.error('Resolve Suitelet URL failed', e);
            return '';
        }
    }

    /* ============================ DATA ============================ */

    function getProviders() {
        var list = [];

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
                list.push({
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

        return list;
    }

    function getDefaultProviderId(providers) {
        for (var i = 0; i < providers.length; i++) {
            if ((providers[i].name || '').toUpperCase().indexOf(DEFAULT_NAME.toUpperCase()) !== -1) {
                return providers[i].id;
            }
        }

        return providers.length ? providers[0].id : '';
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
        var win = getWeekWindow();

        var columns = [
            search.createColumn({
                name: 'internalid',
                summary: "GROUP"
            }),
            search.createColumn({
                name: 'tranid',
                summary: "GROUP"
            }),
            search.createColumn({
                name: 'entity',
                summary: "GROUP"
            }),
            search.createColumn({
                name: DATE_FIELD,
                summary: "MAX"
            }),
            search.createColumn({
                name: 'formulanumericv',
                summary: "SUM",
                formula: "NVL({item.custitem_item_volume},0)*{quantity}"
            }),
            search.createColumn({
                name: 'formulanumericc',
                summary: "SUM",
                formula: "{quantity}"
            }),
        ];

        if (includeZone) {
            columns.push(search.createColumn({
                name: ZONE_FIELD,
                summary: "GROUP"
            }));
        }

        var soSearch = search.create({
            type: search.Type.SALES_ORDER,
            filters: [
                ["mainline","is","F"], 
                "AND", 
                ["taxline","is","F"], 
                "AND", 
                ["shipping","is","F"], 
                'AND',
                [PROVIDER_FIELD, 'anyof', providerId],
                 'AND',
                [DATE_FIELD,"onorafter","today"] 
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

                var rawDate = getResultValue(r, DATE_FIELD, "MAX");
                var shipDate = parseNetSuiteDate(rawDate);

                if (!shipDate) {
                    return;
                }

                var dayNum = shipDate.getDay();

                // Skip weekend expected shipping dates
                if (dayNum === 0 || dayNum === 6) {
                    return;
                }

                var weekInfo = isoWeekInfo(shipDate);
                var internalid = getResultValue(r, 'internalid', 'GROUP') : '';
                var zoneValue = includeZone ? getResultValue(r, ZONE_FIELD, 'GROUP') : '';
                var zoneText = includeZone ? getResultText(r, ZONE_FIELD, 'GROUP') : '';
                var zone = normalizeZone(zoneValue, zoneText);

                rows.push({
                    soId: String(internalid),
                    tranid: String(getResultValue(r, 'tranid', 'GROUP') || ''),
                    customer: String(getResultText(r, 'entity', 'GROUP') || getResultValue(r, 'entity', 'GROUP') || ''),
                    volume: numVal(getResultValue(r, 'formulanumericv', 'SUM')),
                    cases: numVal(getResultValue(r, 'formulanumericc', 'SUM')),
                    ymd: toYMD(shipDate),
                    week: weekInfo.week,
                    weekYear: weekInfo.year,
                    weekKey: weekInfo.key,
                    weekLabel: 'Week ' + weekInfo.week,
                    zone: zone,
                    zoneLabel: zone === '-' ? 'No Zone' : 'Zone ' + zone,
                    dayNum: dayNum,
                    dayLabel: DAY_LABELS[dayNum] || '',
                    dayShort: DAY_SHORT[dayNum] || ''
                });
            });
        });

        // Pallet/Cases body fields are not reliably returned by the search columns,
        // so pull current values per SO via lookupFields (cheap, no full record.load).
        hydrateTotals(rows);

        rows.sort(function (a, b) {
            return cmp(a.weekKey, b.weekKey) ||
                (zoneSortKey(a.zone) - zoneSortKey(b.zone)) ||
                cmp(a.ymd, b.ymd) ||
                cmp(a.tranid, b.tranid);
        });

        return rows;
    }

    function hydrateTotals(rows) {
        if (!rows.length) {
            return;
        }

        rows.forEach(function (row) {
            try {
                var vals = search.lookupFields({
                    type: search.Type.SALES_ORDER,
                    id: row.soId,
                    columns: [VOLUME_FIELD, CASES_FIELD]
                });

                row.volume = numVal(vals[VOLUME_FIELD]);
                row.cases = numVal(vals[CASES_FIELD]);

            } catch (e) {
                log.error('lookupFields failed for SO ' + row.soId, e);
            }
        });
    }

    function getResultValue(result, fieldId, summary) {
        try {
            return result.getValue({
                name: fieldId,
                summary: summary
            });
        } catch (e1) {
            try {
                return result.getValue({name: fieldId, summary: summary });
            } catch (e2) {
                return '';
            }
        }
    }

    function getResultText(result, fieldId, summary) {
        try {
            return result.getText({
                name: fieldId,
                summary: summary
            });
        } catch (e1) {
            try {
                return result.getText({name: fieldId, summary: summary });
            } catch (e2) {
                return '';
            }
        }
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
                if (!edit || !edit.id || !edit.shipdate) {
                    return;
                }

                var newDate = dateFromYMD(edit.shipdate);

                if (!newDate) {
                    throw new Error('Invalid date: ' + edit.shipdate);
                }

                var day = newDate.getDay();

                if (day === 0 || day === 6) {
                    skippedWeekend++;
                    return;
                }

                var values = {};
                values[DATE_FIELD] = newDate;

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

        var msg = 'Saved ' + ok + ' expected ship date(s).';

        if (skippedWeekend) {
            msg += ' ' + skippedWeekend + ' weekend date(s) skipped.';
        }

        if (fail) {
            msg += ' ' + fail + ' failed. See script logs.';
        }

        return msg;
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

        var html = '';

        html += '<div style="font-family:Arial,Helvetica,sans-serif;color:#111827">';
        html += '<h2 style="margin:0 0 6px">TOSAN - Weekly Shipping Schedule</h2>';
        html += '<p style="margin:0 0 14px">Provider: <b>' + esc(providerName) + '</b></p>';

        html += '<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;width:100%;font-size:12px">';
        html += '<thead>';
        html += '<tr style="background:#f3f4f6">';
        html += '<th align="left">Week</th>';
        html += '<th align="left">Zone</th>';
        html += '<th align="left">Day</th>';
        html += '<th align="left">Sales Order</th>';
        html += '<th align="left">Customer</th>';
        html += '<th align="right">Pallet</th>';
        html += '<th align="right">Cases</th>';
        html += '<th align="left">Expected Ship Date</th>';
        html += '</tr>';
        html += '</thead>';
        html += '<tbody>';

        if (!data.length) {
            html += '<tr><td colspan="8" align="center">No sales orders found.</td></tr>';
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
                    html += '<td>' + esc(r.dayLabel) + '</td>';
                    html += '<td>' + esc(r.tranid) + '</td>';
                    html += '<td>' + esc(r.customer) + '</td>';
                    html += '<td align="right">' + fmt(r.volume) + '</td>';
                    html += '<td align="right">' + fmt(r.cases) + '</td>';
                    html += '<td>' + esc(r.ymd) + '</td>';
                    html += '</tr>';
                });
            });
        });

        html += '</tbody>';
        html += '</table>';
        html += '</div>';

        return html;
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
        var z = {};

        data.forEach(function (r) {
            pallet += Number(r.volume) || 0;
            cases += Number(r.cases) || 0;

            if (r.zone && r.zone !== '-') {
                z[r.zone] = true;
            }
        });

        return {
            orders: data.length,
            volume: round2(pallet),
            cases: round2(cases),
            zones: Object.keys(z).length
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

    return {
        onRequest: onRequest
    };
});