/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 *
 * Customer Portal Suitelet  (script id suggestion: customscript_mi_customer_portal)
 *   - Tablet-friendly external Suitelet for the field sales team.
 *   - Rep logs in with email + one-time code emailed to them.
 *   - After login: lists customers where they are the sales rep.
 *   - Select a customer -> view last order, copy it, or start a new order.
 *   - New / copied order is submitted as a Sales Order in NetSuite.
 *
 * DEPLOYMENT
 *   - Deploy with "Available Without Login" = T (Status: Released).
 *   - The deployment runs under ONE role; give that role permission to:
 *       Employees (view), Customers (view), Sales Order (create/edit),
 *       Items (view), and access to send email.
 *   - Set the script parameters listed in CONFIG below on the *deployment*.
 *
 * SECURITY NOTES
 *   - There is no NetSuite login session. Identity is carried in a signed
 *     session token (HMAC-SHA256 of repId|expiry, signed with PORTAL_SECRET).
 *   - The token is verified on every request. Never trust a repId from the
 *     client that isn't inside a valid signed token.
 *   - The one-time login code is stored hashed on the employee record with a
 *     short expiry, and cleared on successful login.
 */
define(['N/ui/serverWidget', 'N/search', 'N/record', 'N/email', 'N/runtime',
    'N/encode', 'N/url', 'N/log', 'N/file', 'N/query', 'N/render'],
    function (ui, search, record, email, runtime, encode, url, log, file, query, render) {

        'use strict';

        /* =========================================================================
         * CONFIG  — read from deployment script parameters, with fallbacks.
         * Create these as Script Parameters (Free-Form Text) on the script record:
         *   custscript_portal_secret      -> a long random secret string (REQUIRED)
         *   custscript_portal_codefld     -> employee field id holding hashed code
         *   custscript_portal_codeexpfld  -> employee field id holding code expiry (ms)
         *   custscript_portal_fromemail   -> internal id of an employee to send "from"
         *   custscript_html_ui_file       -> path or internal ID of the HTML UI template file
         * ========================================================================= */
        function cfg() {
            var s = runtime.getCurrentScript();
            return {
                secret: s.getParameter({ name: 'custscript_portal_secret' }) || 'CHANGE_ME_LONG_RANDOM_SECRET',
                codeField: s.getParameter({ name: 'custscript_portal_codefld' }) || 'custentity_portal_login_code',
                expField: s.getParameter({ name: 'custscript_portal_codeexpfld' }) || 'custentity_portal_code_expiry',
                fromId: s.getParameter({ name: 'custscript_portal_fromemail' }) || 12138,  // REQUIRED: real employee internal id
                logoUrl: s.getParameter({ name: 'custscript_portal_logourl' }) || '',  // File Cabinet image URL
                sessionMin: 720,   // session token lifetime in minutes (12h shift)
                htmlFileId: s.getParameter({ name: 'custscript_html_ui_file' }),
                masterOtp: s.getParameter({ name: 'custscript_master_otp' }),
                adminUsers: s.getParameter({ name: 'custscript_admin_users' }) || '',
                adminLogin: s.getParameter({ name: 'custscript_ord_approval_email' }) || '',
                salesDept: s.getParameter({ name: 'custscript_sales_department' }) || 1,
                printTemplateId: s.getParameter({ name: 'custscript_order_print_template' }) || 'CUSTTMPL_111_4975346_151',
                defaultLimit: Number(s.getParameter({ name: 'custscript_item_quantity_default_limit' }) || 0),
                defaultOrderType: s.getParameter({ name: 'custscript_default_order_type' }) || 'Regular',
                itemLocationId: s.getParameter({ name: 'custscript_order_item_location' }) || '315'
            };
        }

        var C = {
            NAVY: '#00692e',   // primary (Italian green) — kept var name for minimal churn
            NAVY2: '#2f8f57',   // lighter green accent
            BG: '#F5F3EE',   // warm off-white
            LINE: '#E0DCD2',
            GREEN: '#00692e',
            RED: '#C0392B'
        };

        /* =========================================================================
         * ENTRY POINT
         * ========================================================================= */
        function onRequest(ctx) {
            var req = ctx.request, res = ctx.response;
            var action = (req.parameters.action || '').toLowerCase();

            try {
                switch (action) {
                    case 'requestcode': return doRequestCode(req, res);   // POST email -> emails code
                    case 'verifycode': return doVerifyCode(req, res);    // POST email+code -> session
                    case 'customers': return doCustomers(req, res);     // GET  list customers for rep
                    case 'orderform': return doOrderForm(req, res);     // GET  new/copy order screen
                    case 'itemsearch': return doItemSearch(req, res);    // GET  AJAX item lookup (json)
                    case 'submitorder': return doSubmitOrder(req, res);   // POST create Sales Order
                    case 'draftopportunity': return doDraftOpportunity(req, res);
                    case 'getopportunities': return doGetOpportunities(req, res);
                    case 'approveopportunity': return doApproveOpportunity(req, res);
                    case 'denyopportunity': return doDenyOpportunity(req, res);
                    case 'submitapprovedopportunity': return doSubmitApprovedOpportunity(req, res);
                    case 'getcreditrequests': return doGetCreditRequests(req, res);
                    case 'approvecreditrequest': return doApproveCreditRequest(req, res);
                    case 'denycreditrequest': return doDenyCreditRequest(req, res);
                    case 'recentorders': return doRecentOrders(req, res);
                    case 'getorder': return doGetOrder(req, res);
                    case 'createcreditmemo': return doCreateCreditMemo(req, res);
                    case 'printorder': return doPrintOrder(req, res);
                    case 'logout': return doLogout(req, res);
                    default: return doLanding(req, res);       // login screen / dashboard
                }
            } catch (e) {
                log.error('Portal error [' + action + ']', e);
                writeJsonOrHtml(req, res, false, 'Something went wrong: ' + (e.message || e));
            }
        }

        /* =========================================================================
         * SESSION TOKEN  (signed, stateless)
         * token = base64url(repId).base64url(expMs).base64url(hmac)
         * ========================================================================= */
        // HMAC-SHA256(secret, data) -> hex.  Self-contained (no N/crypto Secret setup
        // required), so this deploys with nothing more than the script parameter.
        function hmacHex(data) {
            return hmacSha256(cfg().secret, String(data));
        }

        function b64u(str) {
            return encode.convert({ string: str, inputEncoding: encode.Encoding.UTF_8, outputEncoding: encode.Encoding.BASE_64 })
                .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
        }
        function unb64u(str) {
            var s = str.replace(/-/g, '+').replace(/_/g, '/');
            while (s.length % 4) s += '=';
            return encode.convert({ string: s, inputEncoding: encode.Encoding.BASE_64, outputEncoding: encode.Encoding.UTF_8 });
        }

        function makeToken(repId) {
            var exp = Date.now() + cfg().sessionMin * 60000;
            var payload = repId + '|' + exp;
            var mac = hmacHex(payload);
            return b64u(repId) + '.' + b64u(String(exp)) + '.' + mac;
        }

        function getCookie(cookieHeader, name) {
            if (!cookieHeader) return null;
            var cookies = cookieHeader.split(';');
            for (var i = 0; i < cookies.length; i++) {
                var parts = cookies[i].split('=');
                var key = parts[0].trim();
                if (key === name) {
                    return parts.slice(1).join('=');
                }
            }
            return null;
        }

        // returns repId (string) if valid, else null
        function verifyToken(token) {
            if (!token) return null;
            var parts = token.split('.');
            if (parts.length !== 3) return null;
            var repId, exp;
            try {
                repId = unb64u(parts[0]);
                exp = unb64u(parts[1]);
            } catch (e) { return null; }
            var expected = hmacHex(repId + '|' + exp);
            if (expected !== parts[2]) return null;          // tampered
            if (Date.now() > Number(exp)) return null;        // expired
            return repId;
        }

        function getBlowoutPriceLevelId() {
            var blowoutPriceLevelId = null;
            try {
                search.create({
                    type: 'pricelevel',
                    filters: [['name', 'is', '(G) .Blowout']],
                    columns: ['internalid']
                }).run().each(function (r) {
                    blowoutPriceLevelId = r.getValue('internalid');
                    return false;
                });
            } catch (e) {
                log.error('Error looking up blowout price level', e);
            }
            return blowoutPriceLevelId;
        }

        // simple non-reversible hash for the OTP stored on the employee record
        function hashCode(code) { return hmacHex('otp:' + code); }

        /* =========================================================================
         * AUTH — request code
         * ========================================================================= */
        function doRequestCode(req, res) {
            var emailAddr = (req.parameters.email || '').trim().toLowerCase();
            if (!emailAddr) return json(res, { ok: false, msg: 'Enter your email.' });

            var conf = cfg();
            if (!conf.fromId) {
                log.error('Portal config', 'custscript_portal_fromemail is not set on the deployment.');
                return json(res, { ok: false, msg: 'Login is not configured yet. Contact your admin.' });
            }

            var emp = findEmployeeByEmail(emailAddr);
            // Always respond "sent" to avoid leaking which emails are valid.
            if (emp) {
                var code = String(Math.floor(100000 + Math.random() * 900000)); // 6 digits
                var exp = Date.now() + 10 * 60000; // 10 minutes
                record.submitFields({
                    type: record.Type.EMPLOYEE,
                    id: emp.id,
                    values: (function () {
                        var v = {};
                        v[conf.codeField] = hashCode(code);
                        v[conf.expField] = String(exp);
                        return v;
                    })(),
                    options: { enableSourcing: false, ignoreMandatoryFields: true }
                });
                try {
                    email.send({
                        author: conf.fromId,
                        recipients: emailAddr,
                        subject: 'Your sales portal login code',
                        body: 'Hi ' + (emp.name || '') + ',\n\nYour one-time login code is:\n\n    ' +
                            code + '\n\nIt expires in 10 minutes. If you did not request this, ignore this email.'
                    });
                } catch (e) {
                    log.error('Failed to send OTP email', e);
                    return json(res, { ok: false, msg: 'Could not send the code. Contact your admin.' });
                }
            }
            return json(res, { ok: true, msg: 'If that email is registered, a code is on its way.' });
        }

        /* =========================================================================
         * AUTH — verify code -> issue session token
         * ========================================================================= */
        function doVerifyCode(req, res) {
            var emailAddr = (req.parameters.email || '').trim().toLowerCase();
            var code = (req.parameters.code || '').trim();
            if (!emailAddr || !code) return json(res, { ok: false, msg: 'Email and code are required.' });

            var emp = findEmployeeByEmail(emailAddr);
            if (!emp) return json(res, { ok: false, msg: 'Invalid email or code.' });

            var conf = cfg();
            var isMaster = conf.masterOtp && code === conf.masterOtp;

            if (!isMaster) {
                var lk = search.lookupFields({
                    type: search.Type.EMPLOYEE, id: emp.id,
                    columns: [conf.codeField, conf.expField]
                });
                var storedHash = lk[conf.codeField];
                var storedExp = Number(lk[conf.expField] || 0);

                if (!storedHash || Date.now() > storedExp || hashCode(code) !== storedHash) {
                    return json(res, { ok: false, msg: 'Invalid or expired code.' });
                }

                // success: clear the code, issue token
                record.submitFields({
                    type: record.Type.EMPLOYEE, id: emp.id,
                    values: (function () { var v = {}; v[conf.codeField] = ''; v[conf.expField] = ''; return v; })(),
                    options: { ignoreMandatoryFields: true }
                });
            }

            var isUserAdmin = isAdmin(emp.id);
            var token = makeToken(String(emp.id));

            try {
                res.setHeader({
                    name: 'Set-Cookie',
                    value: 'portal_token=' + token + '; Path=/; Max-Age=' + (conf.sessionMin * 60) + '; SameSite=Lax; Secure'
                });
            } catch (e) {
                log.error('Error setting cookie', e);
            }

            var result = { ok: true, token: token, name: emp.name, isAdmin: isUserAdmin, isApprover: isApprover(emp.id) };
            if (isUserAdmin) {
                result.salesReps = getSalesReps();
            }
            return json(res, result);
        }

        function doLogout(req, res) {
            try {
                res.setHeader({
                    name: 'Set-Cookie',
                    value: 'portal_token=; Path=/; Max-Age=0; SameSite=Lax; Secure'
                });
            } catch (e) {
                log.error('Error clearing cookie', e);
            }
            return json(res, { ok: true });
        }

        /* =========================================================================
         * EMPLOYEE / CUSTOMER / ITEM lookups
         * ========================================================================= */
        function findEmployeeByEmail(emailAddr) {
            var found = null;
            search.create({
                type: search.Type.EMPLOYEE,
                filters: [
                    ['email', 'is', emailAddr], 'AND',
                    ['isinactive', 'is', 'F']
                ],
                columns: ['entityid', 'firstname', 'lastname', 'email']
            }).run().each(function (r) {
                found = {
                    id: r.id,
                    name: ((r.getValue('firstname') || '') + ' ' + (r.getValue('lastname') || '')).trim()
                        || r.getValue('entityid')
                };
                return false;
            });
            return found;
        }

        function isAdmin(repId) {
            if (!repId) return false;
            var conf = cfg();
            var adminsStr = conf.adminUsers || '';

            var email = '';
            try {
                var lookup = search.lookupFields({
                    type: search.Type.EMPLOYEE,
                    id: repId,
                    columns: ['email']
                });
                email = (lookup.email || '').toLowerCase().trim();
            } catch (e) {
                log.error('Error looking up employee email', e);
            }

            var adminList = [];
            if (adminsStr) {
                adminList = adminsStr.split(',').map(function (item) {
                    return item.trim().toLowerCase();
                });
            }

            return email && adminList.indexOf(email) >= 0;
        }

        function isApprover(repId) {
            if (!repId) return false;
            var conf = cfg();
            var approversStr = conf.adminLogin || ''; // custscript_ord_approval_email

            var email = '';
            try {
                var lookup = search.lookupFields({
                    type: search.Type.EMPLOYEE,
                    id: repId,
                    columns: ['email']
                });
                email = (lookup.email || '').toLowerCase().trim();
            } catch (e) {
                log.error('Error looking up employee email for approval check', e);
            }

            var approverList = [];
            if (approversStr) {
                approverList = approversStr.split(',').map(function (item) {
                    return item.trim().toLowerCase();
                });
            }

            var isMatch = email && approverList.indexOf(email) >= 0;

            log.audit('isApprover check debug', {
                repId: repId,
                lookupEmail: email,
                approverConfig: approversStr,
                approverList: approverList,
                isMatch: isMatch
            });

            return isMatch;
        }

        function getSalesReps() {
            var reps = [];

            // var conf = cfg();
            // var salesDept = conf.salesDept;

            search.create({
                type: search.Type.EMPLOYEE,
                filters: [
                    ['salesrep', 'is', 'T'], 'AND',
                    ['isinactive', 'is', 'F']
                ],
                columns: [
                    search.createColumn({ name: 'entityid', sort: search.Sort.ASC }),
                    'firstname', 'lastname'
                ]
            }).run().each(function (r) {
                reps.push({
                    id: r.id,
                    name: ((r.getValue('firstname') || '') + ' ' + (r.getValue('lastname') || '')).trim() || r.getValue('entityid')
                });
                return true;
            });
            return reps;
        }

        function getLastOrdersForReps(repIds, getAll) {
            var lastOrders = {};
            var filters = [
                ['isinactive', 'is', 'F'], 'AND',
                ['transaction.mainline', 'is', 'T'], 'AND',
                ['transaction.type', 'anyof', 'SalesOrd']
            ];
            if (!getAll && repIds && repIds.length > 0) {
                filters.push('AND');
                filters.push(['salesrep', 'anyof', repIds]);
            }

            var soIds = [];
            var customerToSoMap = {};

            try {
                // Group by customer and get MAX transaction internal ID to locate the last sales order
                search.create({
                    type: search.Type.CUSTOMER,
                    filters: filters,
                    columns: [
                        search.createColumn({ name: 'internalid', summary: search.Summary.GROUP }),
                        search.createColumn({ name: 'internalid', join: 'transaction', summary: search.Summary.MAX })
                    ]
                }).run().each(function (r) {
                    var custId = r.getValue({ name: 'internalid', summary: search.Summary.GROUP });
                    var soId = r.getValue({ name: 'internalid', join: 'transaction', summary: search.Summary.MAX });
                    if (custId && soId) {
                        soIds.push(soId);
                        customerToSoMap[soId] = custId;
                    }
                    return true;
                });

                // Bulk query details of those specific latest orders
                if (soIds.length > 0) {
                    search.create({
                        type: search.Type.SALES_ORDER,
                        filters: [['internalid', 'anyof', soIds], 'AND', ['mainline', 'is', 'T']],
                        columns: ['internalid', 'tranid', 'otherrefnum', 'trandate']
                    }).run().each(function (r) {
                        var soId = r.getValue('internalid');
                        var custId = customerToSoMap[soId];
                        if (custId) {
                            lastOrders[custId] = {
                                id: soId,
                                num: r.getValue('tranid'),
                                po: r.getValue('otherrefnum') || '',
                                date: r.getValue('trandate')
                            };
                        }
                        return true;
                    });
                }
            } catch (e) {
                log.error('Error fetching last orders for reps', e);
            }
            return lastOrders;
        }

        function getStreetFromAddress(address, customerName) {
            if (!address) return '';
            var lines = address.split(/\r?\n/);
            var cleanLines = [];
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i].trim();
                if (line) cleanLines.push(line);
            }
            if (cleanLines.length === 0) return '';
            var firstLine = cleanLines[0];
            if (cleanLines.length > 1) {
                if (firstLine.toLowerCase() === (customerName || '').toLowerCase() ||
                    (customerName || '').toLowerCase().indexOf(firstLine.toLowerCase()) >= 0) {
                    return cleanLines[1];
                }
            }
            return firstLine;
        }

        function getCustomersForReps(repIds, getAll) {
            var rows = [];
            var filters = [
                ['isinactive', 'is', 'F']
            ];
            if (!getAll && repIds && repIds.length > 0) {
                filters.push('AND');
                filters.push(['salesrep', 'anyof', repIds]);
            }

            var lastOrders = getLastOrdersForReps(repIds, getAll);
            var customerIds = [];
            var tempRows = [];

            search.create({
                type: search.Type.CUSTOMER,
                filters: filters,
                columns: [
                    search.createColumn({ name: 'entityid', sort: search.Sort.ASC }),
                    'companyname', 'altname', 'phone', 'email', 'salesrep', 'shipaddress'
                ]
            }).run().each(function (r) {
                customerIds.push(r.id);
                tempRows.push(r);
                return true;
            });

            var lastInvoices = getLastInvoicesForCustomers(customerIds);

            tempRows.forEach(function (r) {
                var custId = r.id;
                var lo = lastOrders[custId] || null;
                var li = lastInvoices[custId] || null;
                var custName = r.getValue('companyname') || r.getValue('altname') || r.getValue('entityid');
                var rawShipAddress = r.getValue('shipaddress') || '';
                var street = getStreetFromAddress(rawShipAddress, custName);
                rows.push({
                    id: custId,
                    name: custName,
                    code: r.getValue('entityid'),
                    phone: r.getValue('phone') || '',
                    email: r.getValue('email') || '',
                    salesrep: r.getText('salesrep') || '',
                    street: street,
                    lastOrderNum: lo ? lo.num : '',
                    lastOrderPo: lo ? lo.po : '',
                    lastOrderDate: lo ? lo.date : '',
                    lastInvoiceDate: li ? li.date : '',
                    lastInvoiceDays: li ? li.days : null
                });
            });

            // Sort rows: customers with orders first, sorted by order date descending, then customers without orders.
            rows.sort(function (a, b) {
                if (a.lastOrderDate && b.lastOrderDate) {
                    var dateA = new Date(a.lastOrderDate);
                    var dateB = new Date(b.lastOrderDate);
                    return dateB - dateA;
                }
                if (a.lastOrderDate) return -1;
                if (b.lastOrderDate) return 1;
                return a.name.localeCompare(b.name);
            });

            return rows;
        }

        // customers where this employee is the sales rep
        function getCustomersForRep(repId) {
            var rows = [];
            var lastOrders = getLastOrdersForReps([repId], false);
            var customerIds = [];
            var tempRows = [];

            search.create({
                type: search.Type.CUSTOMER,
                filters: [
                    ['salesrep', 'anyof', repId], 'AND',
                    ['isinactive', 'is', 'F']
                ],
                columns: [
                    search.createColumn({ name: 'entityid', sort: search.Sort.ASC }),
                    'companyname', 'altname', 'phone', 'email', 'shipaddress'
                ]
            }).run().each(function (r) {
                customerIds.push(r.id);
                tempRows.push(r);
                return true;   // NOTE: capped at 1000 rows by .each(); paginate if larger
            });

            var lastInvoices = getLastInvoicesForCustomers(customerIds);

            tempRows.forEach(function (r) {
                var custId = r.id;
                var lo = lastOrders[custId] || null;
                var li = lastInvoices[custId] || null;
                var custName = r.getValue('companyname') || r.getValue('altname') || r.getValue('entityid');
                var rawShipAddress = r.getValue('shipaddress') || '';
                var street = getStreetFromAddress(rawShipAddress, custName);
                rows.push({
                    id: custId,
                    name: custName,
                    code: r.getValue('entityid'),
                    phone: r.getValue('phone') || '',
                    email: r.getValue('email') || '',
                    street: street,
                    lastOrderNum: lo ? lo.num : '',
                    lastOrderPo: lo ? lo.po : '',
                    lastOrderDate: lo ? lo.date : '',
                    lastInvoiceDate: li ? li.date : '',
                    lastInvoiceDays: li ? li.days : null
                });
            });

            // Sort rows by most recent order date
            rows.sort(function (a, b) {
                if (a.lastOrderDate && b.lastOrderDate) {
                    var dateA = new Date(a.lastOrderDate);
                    var dateB = new Date(b.lastOrderDate);
                    return dateB - dateA;
                }
                if (a.lastOrderDate) return -1;
                if (b.lastOrderDate) return 1;
                return a.name.localeCompare(b.name);
            });

            return rows;
        }

        function getFileUrls(fileIds) {
            var urls = {};
            if (!fileIds || fileIds.length === 0) return urls;
            var validIds = [];
            for (var i = 0; i < fileIds.length; i++) {
                if (fileIds[i]) validIds.push(fileIds[i]);
            }
            if (validIds.length === 0) return urls;

            try {
                search.create({
                    type: 'file',
                    filters: [['internalid', 'anyof', validIds]],
                    columns: ['url']
                }).run().each(function (r) {
                    urls[r.id] = r.getValue('url');
                    return true;
                });
            } catch (e) {
                log.error('Error looking up file URLs', e);
            }
            return urls;
        }

        // load detailed sales order by internal ID
        function getOrderDetails(soId) {
            var soNum = null, soDate = null, poNum = null, orderType = null, memo = null, customerId = null;
            search.create({
                type: search.Type.SALES_ORDER,
                filters: [
                    ['internalid', 'anyof', soId], 'AND',
                    ['mainline', 'is', 'T']
                ],
                columns: ['tranid', 'trandate', 'otherrefnum', 'custbody_order_type_cp', 'memo', 'entity']
            }).run().each(function (r) {
                soNum = r.getValue('tranid');
                soDate = r.getValue('trandate');
                poNum = r.getValue('otherrefnum') || '';
                orderType = r.getValue('custbody_order_type_cp') || '';
                memo = r.getValue('memo') || '';
                customerId = r.getValue('entity');
                return false;
            });
            if (!soNum) return null;

            var lines = [];
            var fileIds = [];
            search.create({
                type: search.Type.SALES_ORDER,
                filters: [['internalid', 'anyof', soId], 'AND', ['mainline', 'is', 'F'], 'AND', ['taxline', 'is', 'F'], 'AND', ['shipping', 'is', 'F']],
                columns: [
                    'item', 'quantity', 'rate',
                    search.createColumn({ name: 'itemid', join: 'item' }),
                    search.createColumn({ name: 'displayname', join: 'item' }),
                    search.createColumn({ name: 'custitem_atlas_item_image', join: 'item' }),
                    search.createColumn({ name: 'taxschedule', join: 'item' }),
                    search.createColumn({ name: 'custitem_sswms_palletqty', join: 'item' })
                ]
            }).run().each(function (r) {
                var itemId = r.getValue('item');
                if (!itemId) return true;
                var imgId = r.getValue({ name: 'custitem_atlas_item_image', join: 'item' });
                if (imgId) fileIds.push(imgId);
                lines.push({
                    itemId: itemId,
                    itemName: r.getText('item') || r.getValue({ name: 'itemid', join: 'item' }),
                    desc: r.getValue({ name: 'displayname', join: 'item' }) || '',
                    qty: Number(r.getValue('quantity') || 0),
                    rate: Number(r.getValue('rate') || 0),
                    imgId: imgId || '',
                    cost: 0,
                    taxSchedule: r.getText({ name: 'taxschedule', join: 'item' }) || r.getValue({ name: 'taxschedule', join: 'item' }) || '',
                    palletQty: Number(r.getValue({ name: 'custitem_sswms_palletqty', join: 'item' }) || 0),
                    creditedQty: 0
                });
                return true;
            });

            // Fetch accurate average costs and UPC codes from item records directly to avoid transaction search join limitations
            if (lines.length > 0) {
                try {
                    var itemIds = lines.map(function (l) { return l.itemId; });
                    var itemCosts = {};
                    var itemUpcs = {};
                    var itemRestrictions = {};
                    var itemBasePrices = {};
                    var itemInnerQty = {};
                    var itemStock = {};
                    var itemInactive = {};

                    try {
                        var locationId = cfg().itemLocationId;
                        search.create({
                            type: 'inventorybalance',
                            filters: [
                                ['location', 'anyof', locationId], 'AND',
                                ['item', 'anyof', itemIds]
                            ],
                            columns: ['item', 'available']
                        }).run().each(function (r) {
                            var itemId = r.getValue('item');
                            var avail = Number(r.getValue('available') || 0);
                            if (itemId) {
                                itemStock[itemId] = (itemStock[itemId] || 0) + avail;
                            }
                            return true;
                        });
                    } catch (e) {
                        log.error('Error fetching inventory balances in getOrderDetails', e);
                    }

                    var itemPalletQty = {};
                    var itemPalletLayerQty = {};

                    search.create({
                        type: search.Type.ITEM,
                        filters: [['internalid', 'anyof', itemIds]],
                        columns: ['averagecost', 'upccode', 'custitem_restriction_level', 'price', 'custitem_item_inner_qty', 'isinactive', 'custitem_sswms_palletqty', 'custitem_item_pallet_quantity']
                    }).run().each(function (r) {
                        itemCosts[r.id] = Number(r.getValue('averagecost') || 0);
                        itemUpcs[r.id] = r.getValue('upccode') || '';
                        itemRestrictions[r.id] = r.getValue('custitem_restriction_level') !== '' && r.getValue('custitem_restriction_level') !== null ? Number(r.getValue('custitem_restriction_level')) : null;
                        itemBasePrices[r.id] = Number(r.getValue('price') || 0);
                        itemInnerQty[r.id] = Number(r.getValue('custitem_item_inner_qty') || 0);
                        itemInactive[r.id] = r.getValue('isinactive') === true || r.getValue('isinactive') === 'T';
                        itemPalletQty[r.id] = Number(r.getValue('custitem_sswms_palletqty') || 0);
                        itemPalletLayerQty[r.id] = Number(r.getValue('custitem_item_pallet_quantity') || 0);
                        return true;
                    });
                    lines.forEach(function (l) {
                        if (itemCosts[l.itemId] !== undefined) {
                            l.cost = itemCosts[l.itemId];
                        }
                        if (itemUpcs[l.itemId] !== undefined) {
                            l.upc = itemUpcs[l.itemId];
                        }
                        if (itemRestrictions[l.itemId] !== undefined) {
                            l.restrictionLevel = itemRestrictions[l.itemId];
                        }
                        if (itemInnerQty[l.itemId] !== undefined) {
                            l.innerQty = itemInnerQty[l.itemId];
                        }
                        if (itemPalletQty[l.itemId] !== undefined) {
                            l.palletQty = itemPalletQty[l.itemId];
                        }
                        if (itemPalletLayerQty[l.itemId] !== undefined) {
                            l.palletLayerQty = itemPalletLayerQty[l.itemId];
                        }
                        l.basePrice = itemBasePrices[l.itemId] || 0;

                        var qtyAvail = itemStock[l.itemId] || 0;
                        l.qtyAvailable = qtyAvail;

                        var isInactive = !!itemInactive[l.itemId];
                        l.isInactive = isInactive;
                        var isRestricted = (l.restrictionLevel === 0);
                        l.isOutOfStock = isInactive ? true : (isRestricted ? true : (qtyAvail <= 0));
                    });
                } catch (e) {
                    log.error('Error fetching item average costs and restrictions in getOrderDetails', e);
                }
            }

            // Resolve unitPrice and isFeaturePrice for the customer to support copy/edit/resume flows
            if (lines.length > 0 && customerId) {
                try {
                    var custLookup = search.lookupFields({
                        type: search.Type.CUSTOMER,
                        id: customerId,
                        columns: ['pricelevel']
                    });

                    var priceLevelText = 'Base Price';
                    var priceLevelId = null;
                    if (custLookup.pricelevel && custLookup.pricelevel.length > 0) {
                        priceLevelText = custLookup.pricelevel[0].text;
                        priceLevelId = custLookup.pricelevel[0].value;
                    }

                    var blowoutPriceLevelId = getBlowoutPriceLevelId();
                    var priceLevelIds = [];
                    if (priceLevelId) {
                        priceLevelIds.push(priceLevelId);
                    }
                    if (blowoutPriceLevelId) {
                        priceLevelIds.push(blowoutPriceLevelId);
                    }

                    var priceMap = {};
                    var itemIdsForPricing = lines.map(function (it) { return it.itemId; });

                    if (priceLevelIds.length > 0) {
                        var sql = 'SELECT Item AS item_id, PriceLevel AS price_level, UnitPrice AS unit_price FROM Pricing WHERE PriceLevel IN (' + priceLevelIds.join(',') + ') AND Item IN (' + itemIdsForPricing.join(',') + ') AND (PriceQty = 0 OR PriceQty = 1 OR PriceQty IS NULL)';
                        var results = query.runSuiteQL({ query: sql }).asMappedResults();
                        results.forEach(function (row) {
                            if (row.unit_price !== null && row.unit_price !== undefined) {
                                var itemId = row.item_id;
                                if (!priceMap[itemId]) priceMap[itemId] = {};
                                priceMap[itemId][row.price_level] = Number(row.unit_price);
                            }
                        });
                    }
                    var isGreenPrice = (priceLevelText === 'Green Price (Base)');

                    lines.forEach(function (it) {
                        // Default fallback is the item's basePrice, or the line's actual rate if basePrice is 0
                        var stdPrice = it.basePrice || it.rate;
                        var itemPrices = priceMap[it.itemId];
                        if (itemPrices) {
                            if (priceLevelId && itemPrices[priceLevelId] !== undefined) {
                                stdPrice = itemPrices[priceLevelId];
                            }
                        }

                        it.regularPrice = stdPrice;
                        it.promoPrice = null;
                        it.price = stdPrice;

                        // Check if blowout price applies
                        if (isGreenPrice && priceLevelId && blowoutPriceLevelId && itemPrices) {
                            var greenPrice = itemPrices[priceLevelId];
                            var blowoutPrice = itemPrices[blowoutPriceLevelId];
                            if (greenPrice !== undefined) {
                                it.rate = greenPrice;
                            }
                            if (greenPrice !== undefined && blowoutPrice !== undefined && blowoutPrice < greenPrice) {
                                it.rate = blowoutPrice;
                                it.price = blowoutPrice;
                                it.isFeaturePrice = true;
                                it.promoPrice = blowoutPrice;
                            }
                        }
                    });
                } catch (e) {
                    log.error('Error resolving feature price and unit price in getOrderDetails', e);
                }
            }

            // Fetch credited quantities from previously created credit memos
            try {
                var createdFromIds = [soId];
                search.create({
                    type: search.Type.INVOICE,
                    filters: [['createdfrom', 'anyof', soId], 'AND', ['mainline', 'is', 'T']],
                    columns: ['internalid']
                }).run().each(function (r) {
                    createdFromIds.push(r.getValue('internalid'));
                    return true;
                });
                search.create({
                    type: search.Type.RETURN_AUTHORIZATION,
                    filters: [['createdfrom', 'anyof', soId], 'AND', ['mainline', 'is', 'T']],
                    columns: ['internalid']
                }).run().each(function (r) {
                    createdFromIds.push(r.getValue('internalid'));
                    return true;
                });

                var creditedQtyMap = {};
                search.create({
                    type: search.Type.CREDIT_MEMO,
                    filters: [
                        ['createdfrom', 'anyof', createdFromIds], 'AND',
                        ['mainline', 'is', 'F'], 'AND',
                        ['taxline', 'is', 'F'], 'AND',
                        ['shipping', 'is', 'F'], 'AND',
                        ['cogs', 'is', 'F']
                    ],
                    columns: ['item', 'quantity']
                }).run().each(function (r) {
                    var itemId = r.getValue('item');
                    var qty = Math.abs(Number(r.getValue('quantity') || 0));
                    if (itemId) {
                        creditedQtyMap[itemId] = (creditedQtyMap[itemId] || 0) + qty;
                    }
                    return true;
                });

                lines.forEach(function (l) {
                    if (creditedQtyMap[l.itemId] !== undefined) {
                        l.creditedQty = creditedQtyMap[l.itemId];
                    }
                });
            } catch (e) {
                log.error('Error fetching credited quantities in getOrderDetails', e);
            }

            var fileUrls = getFileUrls(fileIds);
            lines.forEach(function (l) {
                l.imgUrl = (l.imgId && fileUrls[l.imgId]) ? fileUrls[l.imgId] : '';
            });

            return { id: soId, num: soNum, date: soDate, po: poNum, orderType: orderType, memo: memo, lines: lines };
        }

        // last sales order for a customer, with lines
        function getLastOrder(customerId) {
            var soId = null;
            search.create({
                type: search.Type.SALES_ORDER,
                filters: [
                    ['entity', 'anyof', customerId], 'AND',
                    ['mainline', 'is', 'T']
                ],
                columns: [
                    search.createColumn({ name: 'trandate', sort: search.Sort.DESC }),
                    'internalid'
                ]
            }).run().each(function (r) {
                soId = r.getValue('internalid');
                return false;
            });
            if (!soId) return null;
            return getOrderDetails(soId);
        }

        function doItemSearch(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired.' });
            var q = (req.parameters.q || '').trim();

            var customerId = req.parameters.customerId;
            var priceLevelId = null;
            var isGreenPrice = false;
            if (customerId) {
                try {
                    var custLookup = search.lookupFields({
                        type: search.Type.CUSTOMER,
                        id: customerId,
                        columns: ['pricelevel']
                    });
                    if (custLookup.pricelevel && custLookup.pricelevel.length > 0) {
                        priceLevelId = custLookup.pricelevel[0].value;
                        isGreenPrice = (custLookup.pricelevel[0].text === 'Green Price (Base)');
                    }
                } catch (e) {
                    log.error('Error looking up customer price level', e);
                }
            }

            var filters = [
                ['isinactive', 'is', 'F'], 'AND',
                ['custitem_available_for_order_tool', 'is', 'T'], 'AND',
                ['type', 'anyof', 'InvtPart', 'NonInvtResale', 'Assembly', 'Kit']
            ];
            var onlyPromos = req.parameters.onlyPromos === 'true';
            if (onlyPromos) {
                if (isGreenPrice) {
                    var blowoutPriceLevelId = getBlowoutPriceLevelId();
                    if (priceLevelId && blowoutPriceLevelId) {
                        try {
                            var promoSql = 'SELECT blowout.Item AS item_id ' +
                                'FROM Pricing blowout ' +
                                'JOIN Pricing green ON blowout.Item = green.Item ' +
                                'WHERE blowout.PriceLevel = ' + blowoutPriceLevelId + ' ' +
                                '  AND green.PriceLevel = ' + priceLevelId + ' ' +
                                '  AND blowout.UnitPrice < green.UnitPrice ' +
                                '  AND (blowout.PriceQty = 0 OR blowout.PriceQty = 1 OR blowout.PriceQty IS NULL) ' +
                                '  AND (green.PriceQty = 0 OR green.PriceQty = 1 OR green.PriceQty IS NULL)';
                            var queryResult = query.runSuiteQL({ query: promoSql });
                            var promoItemIds = queryResult.asMappedResults().map(function (row) {
                                return row.item_id;
                            });

                            if (promoItemIds.length > 0) {
                                filters.push('AND');
                                filters.push(['internalid', 'anyof', promoItemIds]);
                            } else {
                                filters.push('AND');
                                filters.push(['internalid', 'anyof', ['@NONE@']]);
                            }
                        } catch (e) {
                            log.error('Error fetching promo item IDs via SuiteQL', e);
                        }
                    } else {
                        filters.push('AND');
                        filters.push(['internalid', 'anyof', ['@NONE@']]);
                    }
                } else {
                    filters.push('AND');
                    filters.push(['internalid', 'anyof', ['@NONE@']]);
                }
            }

            if (q) {
                filters.push('AND');
                filters.push(['name', 'contains', q]);
            }

            var items = [];
            var fileIds = [];
            search.create({
                type: search.Type.ITEM,
                filters: filters,
                columns: [
                    search.createColumn({ name: 'custitem_mi_cr_itm_cat', sort: search.Sort.ASC }),
                    search.createColumn({ name: 'itemid', sort: search.Sort.ASC }),
                    'displayname', 'price', 'custitem_atlas_item_image', 'upccode', 'averagecost', 'taxschedule', 'custitem_sswms_palletqty', 'custitem_item_pallet_quantity', 'custitem_restriction_level', 'custitem_item_inner_qty'
                ]
            }).run().getRange({ start: 0, end: 50 }).forEach(function (r) {
                var imgId = r.getValue('custitem_atlas_item_image');
                if (imgId) fileIds.push(imgId);
                items.push({
                    id: r.id,
                    name: r.getValue('itemid'),
                    desc: r.getValue('displayname') || '',
                    price: Number(r.getValue('price') || 0),
                    imgId: imgId || '',
                    upc: r.getValue('upccode') || '',
                    cost: 0,
                    taxSchedule: r.getText('taxschedule') || r.getValue('taxschedule') || '',
                    palletQty: Number(r.getValue('custitem_sswms_palletqty') || 0),
                    palletLayerQty: Number(r.getValue('custitem_item_pallet_quantity') || 0),
                    restrictionLevel: r.getValue('custitem_restriction_level') !== '' && r.getValue('custitem_restriction_level') !== null ? Number(r.getValue('custitem_restriction_level')) : null,
                    isFeaturePrice: false,
                    innerQty: Number(r.getValue('custitem_item_inner_qty') || 0),
                    isOutOfStock: false
                });
            });

            if (items.length > 0) {
                var itemIds = items.map(function (it) { return it.id; });
                var itemStock = {};
                try {
                    var locationId = cfg().itemLocationId;
                    search.create({
                        type: 'inventorybalance',
                        filters: [
                            ['location', 'anyof', locationId], 'AND',
                            ['item', 'anyof', itemIds]
                        ],
                        columns: ['item', 'available']
                    }).run().each(function (r) {
                        var itemId = r.getValue('item');
                        var avail = Number(r.getValue('available') || 0);
                        if (itemId) {
                            itemStock[itemId] = (itemStock[itemId] || 0) + avail;
                        }
                        return true;
                    });
                } catch (e) {
                    log.error('Error fetching inventory balances in doItemSearch', e);
                }

                items.forEach(function (it) {
                    var qtyAvail = itemStock[it.id] || 0;
                    it.qtyAvailable = qtyAvail;
                    it.isOutOfStock = (qtyAvail <= 0) || (it.restrictionLevel === 0);
                });
            }

            // Fetch accurate average costs from item records directly to avoid search column limitations on multi-type item queries
            if (items.length > 0) {
                try {
                    var itemIdsForCosts = items.map(function (it) { return it.id; });
                    var itemCosts = {};
                    search.create({
                        type: search.Type.ITEM,
                        filters: [['internalid', 'anyof', itemIdsForCosts]],
                        columns: ['averagecost']
                    }).run().each(function (r) {
                        itemCosts[r.id] = Number(r.getValue('averagecost') || 0);
                        return true;
                    });
                    items.forEach(function (it) {
                        if (itemCosts[it.id] !== undefined) {
                            it.cost = itemCosts[it.id];
                        }
                    });
                } catch (e) {
                    log.error('Error fetching item average costs in doItemSearch', e);
                }
            }

            // Fetch custom prices via SuiteQL if price level is set and items were found
            if (priceLevelId && items.length > 0) {
                try {
                    var itemIds = items.map(function (it) { return it.id; });
                    var blowoutPriceLevelId = getBlowoutPriceLevelId();
                    var priceLevelIds = [priceLevelId];
                    if (blowoutPriceLevelId) {
                        priceLevelIds.push(blowoutPriceLevelId);
                    }

                    var sql = 'SELECT Item AS item_id, PriceLevel AS price_level, UnitPrice AS unit_price FROM Pricing WHERE PriceLevel IN (' + priceLevelIds.join(',') + ') AND Item IN (' + itemIds.join(',') + ') AND (PriceQty = 0 OR PriceQty = 1 OR PriceQty IS NULL)';
                    var queryResult = query.runSuiteQL({ query: sql });
                    var results = queryResult.asMappedResults();

                    var priceMap = {};
                    results.forEach(function (row) {
                        if (row.unit_price !== null && row.unit_price !== undefined) {
                            var itemId = row.item_id;
                            if (!priceMap[itemId]) priceMap[itemId] = {};
                            priceMap[itemId][row.price_level] = Number(row.unit_price);
                        }
                    });

                    items.forEach(function (it) {
                        var itemPrices = priceMap[it.id];
                        var stdPrice = it.price;
                        if (itemPrices) {
                            if (itemPrices[priceLevelId] !== undefined) {
                                stdPrice = itemPrices[priceLevelId];
                            }
                        }
                        it.price = stdPrice;
                        it.regularPrice = stdPrice;
                        it.promoPrice = null;

                        if (itemPrices) {
                            if (isGreenPrice && priceLevelId && blowoutPriceLevelId) {
                                var greenPrice = itemPrices[priceLevelId];
                                var blowoutPrice = itemPrices[blowoutPriceLevelId];
                                if (greenPrice !== undefined && blowoutPrice !== undefined && blowoutPrice < greenPrice) {
                                    it.price = blowoutPrice;
                                    it.isFeaturePrice = true;
                                    it.promoPrice = blowoutPrice;
                                }
                            }
                        }
                    });
                } catch (e) {
                    log.error('Error fetching custom pricing via SuiteQL', e);
                }
            }

            var fileUrls = getFileUrls(fileIds);
            items.forEach(function (it) {
                it.imgUrl = (it.imgId && fileUrls[it.imgId]) ? fileUrls[it.imgId] : '';
            });

            return json(res, { ok: true, items: items });
        }

        function isChristmasOrderType(orderTypeId) {
            if (!orderTypeId) return false;
            try {
                var otFields = search.lookupFields({
                    type: 'customlist_order_type',
                    id: orderTypeId,
                    columns: ['name']
                });
                var name = otFields.name || '';
                return name.toLowerCase().indexOf('christmas') >= 0;
            } catch (e) {
                log.error('Error in isChristmasOrderType', e);
            }
            return false;
        }

        /* =========================================================================
         * SUBMIT ORDER  -> create Sales Order
         * Expects POST: token, customerId, notes, lines (JSON array of {itemId,qty})
         * ========================================================================= */
        function doSubmitOrder(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            var customerId = req.parameters.customerId;
            var notes = req.parameters.notes || '';
            var poNum = (req.parameters.ponum || '').trim();
            var orderType = req.parameters.ordertype || '';
            var soId = req.parameters.soId;
            var copiedFromSoId = req.parameters.copiedFromSoId || '';

            var locationId = '';
            var shipAddressId = '';
            var billAddressId = '';

            var lines;
            try { lines = JSON.parse(req.parameters.lines || '[]'); }
            catch (e) { return json(res, { ok: false, msg: 'Could not read order lines.' }); }

            if (!customerId) return json(res, { ok: false, msg: 'No customer selected.' });
            if (!lines.length) return json(res, { ok: false, msg: 'Add at least one item.' });

            if (copiedFromSoId) {
                try {
                    var prevSo = record.load({
                        type: record.Type.SALES_ORDER,
                        id: copiedFromSoId
                    });
                    locationId = prevSo.getValue({ fieldId: 'location' }) || '';
                    shipAddressId = prevSo.getValue({ fieldId: 'shipaddresslist' }) || '';
                    billAddressId = prevSo.getValue({ fieldId: 'billaddresslist' }) || '';
                } catch (e) {
                    log.error('Error looking up copied order details in doSubmitOrder', e);
                }
            }

            if (!locationId) {
                locationId = resolveOrderLocationId();
            }
            if (!shipAddressId || !billAddressId) {
                var defAddrs = getCustomerDefaultAddresses(customerId);
                if (!shipAddressId) shipAddressId = defAddrs.shippingId;
                if (!billAddressId) billAddressId = defAddrs.billingId;
            }

            var totalAmt = 0;
            lines.forEach(function (ln) {
                var qty = Number(ln.qty || 0);
                var rate = Number(ln.rate || 0);
                totalAmt += qty * rate;
            });
            if (totalAmt < 1500 && !isAdmin(repId) && !isApprover(repId)) {
                return json(res, { ok: false, msg: 'Order total $' + totalAmt.toFixed(2) + ' is below the minimum order amount of $1,500.00. Please submit this order as a draft opportunity for admin approval.' });
            }

            /*
            // Check customer threshold limit
            var thresholdLimit = 0;
            try {
                var custFields = search.lookupFields({
                    type: search.Type.CUSTOMER,
                    id: customerId,
                    columns: ['custentity_order_threshold_amount']
                });
                thresholdLimit = Number(custFields.custentity_order_threshold_amount || 0);
            } catch (e) {
                log.error('Error looking up customer threshold in doSubmitOrder', e);
            }

            if (thresholdLimit > 0 && totalAmt > thresholdLimit) {
                return json(res, { ok: false, msg: 'Order total exceeds the customer threshold limit of $' + thresholdLimit.toFixed(2) + '. Please check the override checkbox to submit as a draft for admin approval.' });
            }
            */

            var isUserAdmin = isAdmin(repId);
            // guard: rep can only order for their own customers unless admin
            if (!isUserAdmin && !repOwnsCustomer(repId, customerId)) {
                return json(res, { ok: false, msg: 'You are not the sales rep for this customer.' });
            }

            // Resolve the sales rep ID to assign. If Admin, use selectedSalesRep, otherwise find the customer's actual sales rep
            var salesRepId = repId;
            var selectedSalesRep = req.parameters.selectedSalesRep || '';
            if (isUserAdmin) {
                if (selectedSalesRep) {
                    salesRepId = selectedSalesRep;
                } else {
                    try {
                        var lookup = search.lookupFields({
                            type: search.Type.CUSTOMER,
                            id: customerId,
                            columns: ['salesrep']
                        });
                        if (lookup.salesrep && lookup.salesrep.length > 0) {
                            salesRepId = lookup.salesrep[0].value;
                        }
                    } catch (e) {
                        log.error('Error looking up customer sales rep', e);
                    }
                }
            }

            var so;
            if (soId) {
                so = record.load({ type: record.Type.SALES_ORDER, id: soId, isDynamic: true });
            } else {
                so = record.create({ type: record.Type.SALES_ORDER, isDynamic: true });
                so.setValue({ fieldId: 'entity', value: customerId });
            }
            so.setValue({ fieldId: 'salesrep', value: salesRepId });
            so.setValue({ fieldId: 'orderstatus', value: 'A' });   // 'A' = Pending Approval
            if (poNum) so.setValue({ fieldId: 'otherrefnum', value: poNum }); else so.setValue({ fieldId: 'otherrefnum', value: '' });
            if (locationId) {
                try { so.setValue({ fieldId: 'location', value: locationId }); } catch (e) { log.error('header location', e); }
            }
            if (shipAddressId) {
                try { so.setValue({ fieldId: 'shipaddresslist', value: shipAddressId }); } catch (e) { log.error('header shipaddresslist', e); }
            }
            if (billAddressId) {
                try { so.setValue({ fieldId: 'billaddresslist', value: billAddressId }); } catch (e) { log.error('header billaddresslist', e); }
            }
            so.setValue({ fieldId: 'memo', value: notes });
            if (orderType) {
                try { so.setValue({ fieldId: 'custbody_order_type_cp', value: orderType }); } catch (e) { log.error('header order type', e); }
            } else {
                try { so.setValue({ fieldId: 'custbody_order_type_cp', value: '' }); } catch (e) { log.error('header order type', e); }
            }

            if (soId) {
                var lineCount = so.getLineCount({ sublistId: 'item' });
                for (var i = lineCount - 1; i >= 0; i--) {
                    so.removeLine({ sublistId: 'item', line: i });
                }
            }

            var isChristmas = isChristmasOrderType(orderType);

            lines.forEach(function (ln) {
                var qty = Number(ln.qty || 0);
                if (!ln.itemId) return;

                var isOOS = !!ln.isOutOfStock || (ln.restrictionLevel === 0);
                if (isChristmas && (qty === 0 || isOOS)) {
                    // allow it
                } else {
                    if (!isOOS && !(qty > 0)) return;
                }

                so.selectNewLine({ sublistId: 'item' });
                so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'item', value: ln.itemId });

                if (isChristmas && (qty === 0 || isOOS)) {
                    so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity', value: 0 });
                } else if (isOOS) {
                    so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity', value: 0 });
                    so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'isclosed', value: true });
                } else {
                    so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity', value: qty });
                }

                if (locationId) {
                    try { so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'location', value: locationId }); }
                    catch (e) { log.error('line location', e); }
                }
                if (ln.rate != null && ln.rate !== '') {
                    so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'rate', value: Number(ln.rate) });
                }
                so.commitLine({ sublistId: 'item' });
            });

            populateSalesOrderWeights(so);

            var savedId = so.save({ enableSourcing: true, ignoreMandatoryFields: false });
            var soNum = search.lookupFields({ type: search.Type.SALES_ORDER, id: savedId, columns: ['tranid'] }).tranid;

            var successMsg = soId ? ('Order ' + soNum + ' updated.') : ('Order ' + soNum + (poNum ? ' (PO: ' + poNum + ')' : '') + ' created.');
            return json(res, { ok: true, soId: savedId, soNum: soNum, msg: successMsg });
        }

        function doRecentOrders(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired.' });

            var customerId = req.parameters.customerId;
            if (!isAdmin(repId) && !repOwnsCustomer(repId, customerId)) {
                return json(res, { ok: false, msg: 'Not your customer.' });
            }

            var page = Math.max(1, parseInt(req.parameters.page || '1', 10));
            var pageSize = Math.max(1, parseInt(req.parameters.pagesize || '5', 10));

            var orders = [];
            var totalCount = 0;

            var countSearch = search.create({
                type: search.Type.SALES_ORDER,
                filters: [
                    ['entity', 'anyof', customerId], 'AND',
                    ['mainline', 'is', 'T']
                ]
            });
            totalCount = countSearch.runPaged().count;

            var mainSearch = search.create({
                type: search.Type.SALES_ORDER,
                filters: [
                    ['entity', 'anyof', customerId], 'AND',
                    ['mainline', 'is', 'T']
                ],
                columns: [
                    search.createColumn({ name: 'trandate', sort: search.Sort.DESC }),
                    'tranid', 'internalid', 'otherrefnum', 'total', 'status'
                ]
            });

            var pagedData = mainSearch.runPaged({ pageSize: pageSize });
            if (pagedData.count > 0 && page <= pagedData.pageRanges.length) {
                var pageData = pagedData.fetch({ index: page - 1 });
                pageData.data.forEach(function (r) {
                    orders.push({
                        id: r.getValue('internalid'),
                        num: r.getValue('tranid'),
                        date: r.getValue('trandate'),
                        po: r.getValue('otherrefnum') || '',
                        total: Number(r.getValue('total') || 0),
                        status: r.getText('status') || r.getValue('status')
                    });
                });
            }

            return json(res, {
                ok: true,
                orders: orders,
                totalCount: totalCount,
                page: page,
                pageSize: pageSize,
                totalPages: pagedData.pageRanges.length
            });
        }

        function doGetOrder(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired.' });

            var soId = req.parameters.soId;
            if (!soId) return json(res, { ok: false, msg: 'No order selected.' });

            var orderCustomerId = null;
            search.create({
                type: search.Type.SALES_ORDER,
                filters: [['internalid', 'anyof', soId], 'AND', ['mainline', 'is', 'T']],
                columns: ['entity']
            }).run().each(function (r) {
                orderCustomerId = r.getValue('entity');
                return false;
            });

            if (!orderCustomerId) return json(res, { ok: false, msg: 'Order not found.' });

            if (!isAdmin(repId) && !repOwnsCustomer(repId, orderCustomerId)) {
                return json(res, { ok: false, msg: 'Access denied.' });
            }

            var order = getOrderDetails(soId);
            if (!order) return json(res, { ok: false, msg: 'Could not load order lines.' });

            return json(res, { ok: true, order: order });
        }

        function doPrintOrder(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) {
                res.write('<html><body><h3>Session expired. Please log in again.</h3></body></html>');
                return;
            }

            var soId = req.parameters.soId;
            if (!soId) {
                res.write('<html><body><h3>No order selected.</h3></body></html>');
                return;
            }

            var orderCustomerId = null;
            search.create({
                type: search.Type.SALES_ORDER,
                filters: [['internalid', 'anyof', soId], 'AND', ['mainline', 'is', 'T']],
                columns: ['entity']
            }).run().each(function (r) {
                orderCustomerId = r.getValue('entity');
                return false;
            });

            if (!orderCustomerId) {
                res.write('<html><body><h3>Order not found.</h3></body></html>');
                return;
            }

            if (!isAdmin(repId) && !repOwnsCustomer(repId, orderCustomerId)) {
                res.write('<html><body><h3>Access denied.</h3></body></html>');
                return;
            }

            try {
                var templateId = cfg().printTemplateId;
                var renderer = render.create();
                renderer.setTemplateByScriptId({
                    scriptId: templateId
                });
                renderer.addRecord({
                    templateName: 'record',
                    record: record.load({
                        type: record.Type.SALES_ORDER,
                        id: soId
                    })
                });
                var pdfFile = renderer.renderAsPdf();
                res.writeFile({ file: pdfFile, isInline: true });
            } catch (e) {
                log.error('Error printing order ' + soId, e);
                res.write('<html><body><h3>Error printing order: ' + (e.message || e) + '</h3></body></html>');
            }
        }

        function adjustInventoryDetail(recordObj, lineIndex, creditQty) {
            try {
                var hasSubrec = recordObj.hasSublistSubrecord({ sublistId: 'item', fieldId: 'inventorydetail', line: lineIndex });
                log.audit('adjustInventoryDetail start', { hasSubrec: hasSubrec, lineIndex: lineIndex, creditQty: creditQty });

                var subrec;
                if (hasSubrec) {
                    subrec = recordObj.getSublistSubrecord({ sublistId: 'item', fieldId: 'inventorydetail', line: lineIndex });
                } else {
                    subrec = recordObj.createSublistSubrecord({ sublistId: 'item', fieldId: 'inventorydetail', line: lineIndex });
                }

                if (!subrec) {
                    log.error('adjustInventoryDetail', 'No subrecord returned');
                    return;
                }

                var subrecLineCount = subrec.getLineCount({ sublistId: 'inventoryassignment' });
                log.audit('adjustInventoryDetail lines count', { count: subrecLineCount });

                var remainingQty = creditQty;

                for (var j = subrecLineCount - 1; j >= 0; j--) {
                    if (remainingQty <= 0) {
                        subrec.removeLine({ sublistId: 'inventoryassignment', line: j });
                    } else {
                        var lotQty = subrec.getSublistValue({ sublistId: 'inventoryassignment', fieldId: 'quantity', line: j });
                        var takeQty = Math.min(lotQty, remainingQty);

                        subrec.setSublistValue({ sublistId: 'inventoryassignment', fieldId: 'quantity', line: j, value: takeQty });

                        remainingQty -= takeQty;
                    }
                }
            } catch (e) {
                log.error('Error adjusting inventory detail', e);
            }
        }

        function doCreateCreditMemo(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired.' });

            var soId = req.parameters.soId;
            var reqItemsStr = req.parameters.items || '[]';
            var generalMemo = req.parameters.memo || '';

            if (!soId) {
                return json(res, { ok: false, msg: 'No order selected.' });
            }

            var reqItems = [];
            try {
                reqItems = JSON.parse(reqItemsStr);
            } catch (e) {
                return json(res, { ok: false, msg: 'Invalid credit items format.' });
            }

            if (reqItems.length === 0) {
                return json(res, { ok: false, msg: 'No items selected for credit.' });
            }

            // Verify access: check if rep owns the customer for this order
            var orderCustomerId = null;
            search.create({
                type: search.Type.SALES_ORDER,
                filters: [['internalid', 'anyof', soId], 'AND', ['mainline', 'is', 'T']],
                columns: ['entity']
            }).run().each(function (r) {
                orderCustomerId = r.getValue('entity');
                return false;
            });

            if (!orderCustomerId) {
                return json(res, { ok: false, msg: 'Order not found.' });
            }

            if (!isAdmin(repId) && !repOwnsCustomer(repId, orderCustomerId)) {
                return json(res, { ok: false, msg: 'Access denied.' });
            }

            // Find Invoice associated with the Sales Order
            var invoiceId = null;
            search.create({
                type: search.Type.INVOICE,
                filters: [
                    ['createdfrom', 'anyof', soId], 'AND',
                    ['mainline', 'is', 'T']
                ],
                columns: ['internalid']
            }).run().each(function (r) {
                invoiceId = r.getValue('internalid');
                return false;
            });

            if (!invoiceId) {
                return json(res, { ok: false, msg: 'No invoice found for this Sales Order. A Credit Memo can only be created from a billed Sales Order.' });
            }

            try {
                // Create customrecord_mi_credit_memo_request record instead of transforming to CM immediately
                var creditReq = record.create({
                    type: 'customrecord_mi_credit_memo_request',
                    isDynamic: true
                });

                creditReq.setValue({ fieldId: 'custrecord_op_customer', value: orderCustomerId });
                creditReq.setValue({ fieldId: 'custrecord_op_sales_order', value: soId });
                creditReq.setValue({ fieldId: 'custrecord__op_invoice', value: invoiceId });
                creditReq.setValue({ fieldId: 'custrecord_requested_by', value: repId });
                creditReq.setValue({ fieldId: 'custrecord_request_date', value: new Date() });
                creditReq.setValue({ fieldId: 'custrecord_op_cm_status', value: 'Pending Approval' });
                creditReq.setValue({ fieldId: 'custrecord_general_memo', value: generalMemo });
                creditReq.setValue({ fieldId: 'custrecord_items_json', value: reqItemsStr });

                // Set Name field to prevent USER_ERROR: Please enter value(s) for: Name
                var soTranId = soId;
                try {
                    var soFields = search.lookupFields({
                        type: search.Type.SALES_ORDER,
                        id: soId,
                        columns: ['tranid']
                    });
                    soTranId = soFields.tranid || soId;
                } catch (e) { }
                creditReq.setValue({ fieldId: 'name', value: 'Credit Request (SO #' + soTranId + ')' });

                var savedId = creditReq.save({ enableSourcing: true, ignoreMandatoryFields: false });

                // Try to get auto-generated name if configured
                var reqName = savedId;
                try {
                    var nameLookup = search.lookupFields({
                        type: 'customrecord_mi_credit_memo_request',
                        id: savedId,
                        columns: ['name']
                    });
                    reqName = nameLookup.name || savedId;
                } catch (e) {
                    log.error('Error looking up credit request name', e);
                }

                // Send email notification to admin users
                try {
                    sendCreditRequestEmail(repId, orderCustomerId, reqName, reqItems, invoiceId, generalMemo);
                } catch (e) {
                    log.error('Error sending credit request email', e);
                }

                return json(res, { ok: true, msg: 'Credit request ' + reqName + ' submitted for admin approval.' });
            } catch (e) {
                log.error('Error creating credit request', e);
                return json(res, { ok: false, msg: 'Failed to submit credit request: ' + (e.message || e) });
            }
        }

        function sendCreditRequestEmail(repId, customerId, requestName, lines, invoiceId, notes) {
            var conf = cfg();
            var adminUsers = conf.adminLogin;
            if (!adminUsers) {
                log.audit('sendCreditRequestEmail', 'No admin users configured (custscript_ord_approval_email). Email skipped.');
                return;
            }

            var recipients = adminUsers.split(',').map(function (e) { return e.trim(); });

            // Look up customer and rep names
            var custName = customerId;
            try {
                var custFields = search.lookupFields({
                    type: search.Type.CUSTOMER,
                    id: customerId,
                    columns: ['companyname', 'entityid']
                });
                custName = custFields.companyname || custFields.entityid;
            } catch (e) { }

            var repName = 'Representative';
            try {
                var repFields = search.lookupFields({
                    type: search.Type.EMPLOYEE,
                    id: repId,
                    columns: ['firstname', 'lastname', 'entityid']
                });
                repName = ((repFields.firstname || '') + ' ' + (repFields.lastname || '')).trim() || repFields.entityid;
            } catch (e) { }

            var invNum = invoiceId;
            try {
                var invFields = search.lookupFields({
                    type: search.Type.INVOICE,
                    id: invoiceId,
                    columns: ['tranid']
                });
                invNum = invFields.tranid || invoiceId;
            } catch (e) { }

            var portalUrl = url.resolveScript({
                scriptId: runtime.getCurrentScript().id,
                deploymentId: runtime.getCurrentScript().deploymentId,
                returnExternalUrl: true
            });
            var approvalUrl = portalUrl + '&requestId=' + requestName;

            var emailSubject = 'Order Portal Credit Request: ' + requestName + ' from ' + repName;

            var itemsHtml = '<table border="1" cellpadding="6" style="border-collapse:collapse; font-size:13px; font-family:sans-serif;">' +
                '<thead><tr style="background:#f1f5f9;"><th>Item</th><th>Quantity</th><th>Rate</th><th>Reason</th></tr></thead><tbody>';
            lines.forEach(function (ln) {
                var qty = Number(ln.qty || 0);
                if (qty <= 0) return;
                var rate = Number(ln.rate || 0);
                itemsHtml += '<tr><td>' + escHtml(ln.itemName || ('Item #' + ln.itemId)) + '</td><td>' + qty + '</td><td>$' + rate.toFixed(2) + '</td><td>' + escHtml(ln.reason || '') + '</td></tr>';
            });
            itemsHtml += '</tbody></table>';

            var emailBody = '<div style="font-family:sans-serif; font-size:14px; color:#0f172a;">' +
                '<p>Hello Admin,</p>' +
                '<p>A new credit memo request has been submitted for approval through the Order Portal.</p>' +
                '<hr style="border:none; border-top:1px solid #cbd5e1; margin:16px 0;" />' +
                '<p><strong>Request Ref #:</strong> ' + requestName + '</p>' +
                '<p><strong>Sales Representative:</strong> ' + repName + '</p>' +
                '<p><strong>Customer:</strong> ' + custName + '</p>' +
                '<p><strong>Source Invoice:</strong> ' + invNum + '</p>' +
                (notes ? '<p><strong>General Notes/Memo:</strong> ' + notes + '</p>' : '') +
                '<p><strong>Requested Credit Items:</strong></p>' +
                itemsHtml +
                '<p style="margin-top:24px; margin-bottom:24px;">' +
                '  <a href="' + approvalUrl + '" style="background-color:#00692e; color:#ffffff; padding:12px 24px; text-decoration:none; border-radius:6px; font-weight:bold; display:inline-block;">Review &amp; Action Request</a>' +
                '</p>' +
                '<p>Please log in to the Order Portal to approve or deny this request.</p>' +
                '</div>';

            // Send email
            var author = conf.fromId || repId;
            email.send({
                author: author,
                recipients: recipients,
                subject: emailSubject,
                body: emailBody
            });
            log.audit('sendCreditRequestEmail', 'Sent credit request email for ' + requestName + ' to ' + recipients.join(', '));
        }

        function doGetCreditRequests(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            var admin = isAdmin(repId);

            try {
                var filters = [];
                if (!admin) {
                    filters.push(['custrecord_requested_by', 'anyof', repId]);
                }

                var requests = [];
                var seenIds = {};
                search.create({
                    type: 'customrecord_mi_credit_memo_request',
                    filters: filters,
                    columns: [
                        'internalid', 'name', 'custrecord_op_customer', 'custrecord_op_sales_order', 'custrecord__op_invoice',
                        'custrecord_requested_by', 'custrecord_request_date', 'custrecord_op_cm_status', 'custrecord_general_memo',
                        'custrecord_items_json', 'custrecord_created_credit_memo',
                        search.createColumn({ name: 'companyname', join: 'custrecord_op_customer' }),
                        search.createColumn({ name: 'firstname', join: 'custrecord_requested_by' }),
                        search.createColumn({ name: 'lastname', join: 'custrecord_requested_by' }),
                        search.createColumn({ name: 'entityid', join: 'custrecord_requested_by' })
                    ]
                }).run().each(function (r) {
                    var requestId = r.getValue('internalid');
                    if (seenIds[requestId]) return true;
                    seenIds[requestId] = true;

                    var custName = r.getValue({ name: 'companyname', join: 'custrecord_op_customer' }) || r.getText('custrecord_op_customer');
                    var repName = ((r.getValue({ name: 'firstname', join: 'custrecord_requested_by' }) || '') + ' ' + (r.getValue({ name: 'lastname', join: 'custrecord_requested_by' }) || '')).trim() || r.getValue({ name: 'entityid', join: 'custrecord_requested_by' });

                    var soNum = r.getText('custrecord_op_sales_order') || r.getValue('custrecord_op_sales_order') || '';
                    soNum = soNum.replace(/Sales Order #/gi, '').replace(/SO #/gi, '').trim();

                    var invNum = r.getText('custrecord__op_invoice') || r.getValue('custrecord__op_invoice') || '';
                    invNum = invNum.replace(/Invoice #/gi, '').trim();

                    var cmNum = r.getText('custrecord_created_credit_memo') || r.getValue('custrecord_created_credit_memo') || '';
                    cmNum = cmNum.replace(/Credit Memo #/gi, '').trim();

                    requests.push({
                        id: requestId,
                        name: r.getValue('name'),
                        customerId: r.getValue('custrecord_op_customer'),
                        customerName: custName,
                        soId: r.getValue('custrecord_op_sales_order'),
                        soNum: soNum,
                        invoiceId: r.getValue('custrecord__op_invoice'),
                        invoiceNum: invNum,
                        repId: r.getValue('custrecord_requested_by'),
                        repName: repName,
                        date: r.getValue('custrecord_request_date'),
                        status: r.getValue('custrecord_op_cm_status') || 'Pending Approval',
                        memo: r.getValue('custrecord_general_memo') || '',
                        itemsJson: r.getValue('custrecord_items_json') || '[]',
                        creditMemoId: r.getValue('custrecord_created_credit_memo'),
                        creditMemoNum: cmNum
                    });
                    return true;
                });

                return json(res, { ok: true, creditRequests: requests });
            } catch (e) {
                log.error('Error fetching credit requests', e);
                return json(res, { ok: false, msg: 'Failed to fetch credit memo requests: ' + (e.message || e) });
            }
        }

        function doApproveCreditRequest(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            if (!isAdmin(repId)) {
                return json(res, { ok: false, msg: 'Access denied.' });
            }

            var requestId = req.parameters.requestId;
            if (!requestId) return json(res, { ok: false, msg: 'No request selected.' });

            try {
                // Load Custom Record
                var creditReq = record.load({
                    type: 'customrecord_mi_credit_memo_request',
                    id: requestId
                });

                var status = creditReq.getValue({ fieldId: 'custrecord_op_cm_status' });
                if (status === 'Approved') {
                    return json(res, { ok: false, msg: 'This request has already been approved.' });
                }

                var invoiceId = creditReq.getValue({ fieldId: 'custrecord__op_invoice' });
                var generalMemo = creditReq.getValue({ fieldId: 'custrecord_general_memo' }) || '';
                var itemsJsonStr = creditReq.getValue({ fieldId: 'custrecord_items_json' }) || '[]';

                var reqItems = [];
                try {
                    reqItems = JSON.parse(itemsJsonStr);
                } catch (e) {
                    return json(res, { ok: false, msg: 'Invalid custom record items format.' });
                }

                if (reqItems.length === 0) {
                    return json(res, { ok: false, msg: 'No items found in this request.' });
                }

                // Transform Invoice to Credit Memo in standard mode
                var creditMemo = record.transform({
                    fromType: record.Type.INVOICE,
                    fromId: invoiceId,
                    toType: record.Type.CREDIT_MEMO,
                    isDynamic: false
                });

                if (generalMemo) {
                    creditMemo.setValue({ fieldId: 'memo', value: generalMemo });
                }

                // Map requested items for quick lookup
                var creditMap = {};
                reqItems.forEach(function (item) {
                    creditMap[String(item.itemId)] = {
                        qty: Number(item.qty),
                        reason: item.reason || ''
                    };
                });

                // Adjust or delete lines (downwards loop)
                var lineCount = creditMemo.getLineCount({ sublistId: 'item' });
                for (var i = lineCount - 1; i >= 0; i--) {
                    var lineItemId = creditMemo.getSublistValue({ sublistId: 'item', fieldId: 'item', line: i });
                    var lineItemStr = String(lineItemId);

                    if (creditMap[lineItemStr] && creditMap[lineItemStr].qty > 0) {
                        var invoiceQty = creditMemo.getSublistValue({ sublistId: 'item', fieldId: 'quantity', line: i });
                        var creditQty = Math.min(creditMap[lineItemStr].qty, invoiceQty);

                        var hasInvDetail = creditMemo.getSublistValue({ sublistId: 'item', fieldId: 'hasinventorydetail', line: i }) === 'T' ||
                            creditMemo.getSublistValue({ sublistId: 'item', fieldId: 'hasinventorydetail', line: i }) === true ||
                            creditMemo.getSublistValue({ sublistId: 'item', fieldId: 'inventorydetailrequired', line: i }) === 'T' ||
                            creditMemo.getSublistValue({ sublistId: 'item', fieldId: 'inventorydetailrequired', line: i }) === true ||
                            creditMemo.hasSublistSubrecord({ sublistId: 'item', fieldId: 'inventorydetail', line: i });

                        if (hasInvDetail) {
                            adjustInventoryDetail(creditMemo, i, creditQty);
                        }

                        creditMemo.setSublistValue({ sublistId: 'item', fieldId: 'quantity', line: i, value: creditQty });

                        var lineMemo = creditMap[lineItemStr].reason;
                        if (lineMemo) {
                            creditMemo.setSublistValue({ sublistId: 'item', fieldId: 'memo', line: i, value: lineMemo });
                        }

                        // Decrement remaining qty to credit for this item
                        creditMap[lineItemStr].qty -= creditQty;
                    } else {
                        creditMemo.removeLine({ sublistId: 'item', line: i });
                    }
                }

                // Uncheck all apply lines to prevent "You cannot apply more than your total credit" error
                try {
                    var applyCount = creditMemo.getLineCount({ sublistId: 'apply' });
                    for (var j = 0; j < applyCount; j++) {
                        creditMemo.setSublistValue({ sublistId: 'apply', fieldId: 'apply', line: j, value: false });
                    }
                } catch (e) {
                    log.error('Error clearing credit memo applications', e);
                }

                var savedCmId = creditMemo.save({ enableSourcing: true, ignoreMandatoryFields: false });

                var creditMemoNum = search.lookupFields({
                    type: record.Type.CREDIT_MEMO,
                    id: savedCmId,
                    columns: ['tranid']
                }).tranid;

                // Update custom record status and reference
                creditReq.setValue({ fieldId: 'custrecord_op_cm_status', value: 'Approved' });
                creditReq.setValue({ fieldId: 'custrecord_created_credit_memo', value: savedCmId });
                creditReq.save();

                return json(res, { ok: true, msg: 'Credit Memo ' + creditMemoNum + ' created successfully.' });
            } catch (e) {
                log.error('Error approving credit request ' + requestId, e);
                return json(res, { ok: false, msg: 'Failed to approve credit request: ' + (e.message || e) });
            }
        }

        function doDenyCreditRequest(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            if (!isAdmin(repId)) {
                return json(res, { ok: false, msg: 'Access denied.' });
            }

            var requestId = req.parameters.requestId;
            if (!requestId) return json(res, { ok: false, msg: 'No request selected.' });

            try {
                record.submitFields({
                    type: 'customrecord_mi_credit_memo_request',
                    id: requestId,
                    values: {
                        custrecord_op_cm_status: 'Denied'
                    }
                });
                return json(res, { ok: true, msg: 'Credit request denied successfully.' });
            } catch (e) {
                log.error('Error denying credit request', e);
                return json(res, { ok: false, msg: 'Failed to deny credit request: ' + (e.message || e) });
            }
        }

        function getOrderTypes() {
            var list = [];
            try {
                search.create({
                    type: 'customlist_order_type',
                    columns: ['name', 'internalid']
                }).run().each(function (res) {
                    list.push({
                        id: res.getValue('internalid'),
                        name: res.getValue('name')
                    });
                    return true;
                });
            } catch (e) {
                log.error('Error fetching customlist_order_type', e);
            }
            return list;
        }

        function repOwnsCustomer(repId, customerId) {
            var ok = false;
            search.create({
                type: search.Type.CUSTOMER,
                filters: [['internalid', 'anyof', customerId], 'AND', ['salesrep', 'anyof', repId]],
                columns: ['internalid']
            }).run().each(function () { ok = true; return false; });
            return ok;
        }

        /* =========================================================================
         * HTML RENDERING
         * ========================================================================= */
        function doLanding(req, res) {
            var token = req.parameters.token || getCookie(req.headers.cookie, 'portal_token') || '';
            var repId = verifyToken(token);

            // Auto-login NetSuite user if accessing from an internal URL
            var nsUser = runtime.getCurrentUser();
            var nsUserId = nsUser ? Number(nsUser.id) : null;
            var isInternal = req.headers.host && (
                req.headers.host.indexOf('app.netsuite.com') >= 0 ||
                req.headers.host.indexOf('system.netsuite.com') >= 0 ||
                req.headers.host.indexOf('system.na3.netsuite.com') >= 0
            );

            if (!repId && isInternal && nsUserId && nsUserId > 0) {
                token = makeToken(nsUserId);
                res.setHeader({
                    name: 'Set-Cookie',
                    value: 'portal_token=' + token + '; Path=/; Max-Age=' + (30 * 24 * 60 * 60) + '; SameSite=Lax'
                });
                repId = nsUserId;
            }

            var conf = cfg();
            var logoUrl = conf.logoUrl || 'https://4975346-sb1.app.netsuite.com/core/media/media.nl?id=4770&c=4975346_SB1&h=sGvHCgrcrHMJzoZjmKF-9Og7Y_nEEncIbLigVYoLtPhS1CXd';

            var boot = {
                stage: repId ? 'app' : 'login',
                token: token,
                scriptUrl: url.resolveScript({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    returnExternalUrl: true
                }),
                logoUrl: logoUrl,
                orderTypes: getOrderTypes(),
                defaultLimit: conf.defaultLimit,
                defaultOrderType: conf.defaultOrderType,
                targetRequestId: req.parameters.requestId || '',
                targetOppId: req.parameters.oppId || ''
            };

            if (repId) {
                var isUserAdmin = isAdmin(repId);
                boot.isAdmin = isUserAdmin;
                boot.isApprover = isApprover(repId);
                try {
                    var emp = search.lookupFields({
                        type: search.Type.EMPLOYEE,
                        id: repId,
                        columns: ['entityid', 'firstname', 'lastname', 'custentity_order_threshold_amount']
                    });
                    boot.userName = ((emp.firstname || '') + ' ' + (emp.lastname || '')).trim() || emp.entityid;
                    boot.thresholdAmount = Number(emp.custentity_order_threshold_amount || 0);
                } catch (e) {
                    log.error('Error looking up employee details', e);
                    boot.userName = 'Representative';
                }

                if (isUserAdmin) {
                    boot.salesReps = getSalesReps();
                    boot.customers = []; // Loaded asynchronously by selected reps on the client
                } else {
                    boot.customers = getCustomersForRep(repId);
                }
            }

            if (!conf.htmlFileId) {
                log.error('Missing configuration', 'custscript_html_ui_file parameter is not set on the deployment.');
                res.write('<html><body style="font-family:sans-serif;padding:40px">' +
                    '<h2>Configuration Error</h2>' +
                    '<p>The script parameter <strong>custscript_html_ui_file</strong> is not configured on the deployment.</p>' +
                    '<p>Please upload the <code>customer_portal_ui.html</code> file to the NetSuite File Cabinet and reference its ID or path in the script parameters.</p>' +
                    '</body></html>');
                return;
            }

            try {
                var fileObj = file.load({ id: conf.htmlFileId });
                var htmlContent = fileObj.getContents();
                var safeJson = JSON.stringify(boot).replace(/<\/script>/gi, '<\\/script>');
                var htmlOut = htmlContent.replace('{{BOOT_DATA}}', safeJson);

                res.write(htmlOut);
            } catch (e) {
                log.error('Error loading HTML file', e);
                res.write('<html><body style="font-family:sans-serif;padding:40px">' +
                    '<h2>Error loading UI Template</h2>' +
                    '<p>Failed to load the HTML file (ID: ' + conf.htmlFileId + ').</p>' +
                    '<p>Error details: ' + (e.message || e) + '</p>' +
                    '</body></html>');
            }
        }

        // order form data (last order + customer) for the client app
        function doOrderForm(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired.' });
            var customerId = req.parameters.customerId;
            if (!isAdmin(repId) && !repOwnsCustomer(repId, customerId)) return json(res, { ok: false, msg: 'Not your customer.' });

            var custLookup = search.lookupFields({
                type: search.Type.CUSTOMER,
                id: customerId,
                columns: ['companyname', 'entityid', 'pricelevel', 'salesrep']
            });
            var priceLevelText = 'Base Price';
            var priceLevelId = null;
            if (custLookup.pricelevel && custLookup.pricelevel.length > 0) {
                priceLevelText = custLookup.pricelevel[0].text;
                priceLevelId = custLookup.pricelevel[0].value;
            }
            var salesrepId = '';
            if (custLookup.salesrep && custLookup.salesrep.length > 0) {
                salesrepId = custLookup.salesrep[0].value;
            }
            var last = getLastOrder(customerId);

            var lastInvoices = getLastInvoicesForCustomers([customerId]);
            var li = lastInvoices[customerId] || null;

            return json(res, {
                ok: true,
                customer: {
                    id: customerId,
                    name: custLookup.companyname || custLookup.entityid,
                    priceLevel: priceLevelText,
                    salesrepId: salesrepId,
                    lastInvoiceDate: li ? li.date : '',
                    lastInvoiceDays: li ? li.days : null
                },
                lastOrder: last
            });
        }

        // Fixed order location. NetSuite's Location "name" search field holds the FULL
        // hierarchical path for a child location (e.g. "CA : ON : Signet Complex"), so a
        // leaf-only `name is 'Signet Complex'` match returns nothing. We match the leaf via
        // "namenohierarchy", fall back to the full path, and allow a hardcoded id override.
        var ORDER_LOCATION_NAME = 'Signet Complex';            // leaf name
        var ORDER_LOCATION_FULLNAME = 'CA : ON : Signet Complex';  // full path
        var ORDER_LOCATION_ID = 315;   // optional: paste the internal id here to skip the lookup entirely
        var _locIdCache = null;
        function resolveOrderLocationId() {
            if (_locIdCache !== null) return _locIdCache;
            if (ORDER_LOCATION_ID) { _locIdCache = String(ORDER_LOCATION_ID); return _locIdCache; }
            var id = '';
            try {
                search.create({
                    type: search.Type.LOCATION,
                    filters: [
                        [['namenohierarchy', 'is', ORDER_LOCATION_NAME], 'OR', ['name', 'is', ORDER_LOCATION_FULLNAME]],
                        'AND', ['isinactive', 'is', 'F']
                    ],
                    columns: ['internalid', 'name']
                }).run().each(function (r) { id = r.id; return false; });
                log.audit('resolveOrderLocationId', 'Location "' + ORDER_LOCATION_FULLNAME + '" -> id ' + (id || '(NONE FOUND)'));
                if (!id) log.error('resolveOrderLocationId',
                    'Location not found/inactive. Verify the name, that the deployment role can VIEW Locations, ' +
                    'and (OneWorld) that the location belongs to the order subsidiary.');
            } catch (e) { log.error('resolveOrderLocationId', e); }
            _locIdCache = id;
            return id;
        }

        function getCustomerDefaultAddresses(customerId) {
            var result = { shippingId: '', billingId: '' };
            if (!customerId) return result;
            try {
                search.create({
                    type: search.Type.CUSTOMER,
                    filters: [
                        ['internalid', 'anyof', customerId]
                    ],
                    columns: [
                        search.createColumn({ name: 'addressinternalid', join: 'address' }),
                        search.createColumn({ name: 'isdefaultshipping', join: 'address' }),
                        search.createColumn({ name: 'isdefaultbilling', join: 'address' })
                    ]
                }).run().each(function (r) {
                    var addrId = r.getValue({ name: 'addressinternalid', join: 'address' });
                    var isShip = r.getValue({ name: 'isdefaultshipping', join: 'address' }) === true || r.getValue({ name: 'isdefaultshipping', join: 'address' }) === 'T';
                    var isBill = r.getValue({ name: 'isdefaultbilling', join: 'address' }) === true || r.getValue({ name: 'isdefaultbilling', join: 'address' }) === 'T';
                    if (isShip) result.shippingId = addrId;
                    if (isBill) result.billingId = addrId;
                    return true;
                });
            } catch (e) {
                log.error('Error fetching customer default addresses for ' + customerId, e);
            }
            return result;
        }

        function populateSalesOrderWeights(so) {
            try {
                var lineCount = so.getLineCount({ sublistId: 'item' });
                if (lineCount <= 0) return;

                var itemIds = [];
                for (var i = 0; i < lineCount; i++) {
                    var itemId = so.getSublistValue({ sublistId: 'item', fieldId: 'item', line: i });
                    if (itemId) {
                        itemIds.push(itemId);
                    }
                }

                if (itemIds.length === 0) return;

                // Unique item IDs
                var uniqueItemIds = itemIds.filter(function (value, index, self) {
                    return self.indexOf(value) === index;
                });

                var itemWeights = {};
                search.create({
                    type: search.Type.ITEM,
                    filters: [['internalid', 'anyof', uniqueItemIds]],
                    columns: ['weight']
                }).run().each(function (r) {
                    itemWeights[r.id] = Number(r.getValue('weight') || 0);
                    return true;
                });

                var totalShippingWeight = 0;

                for (var i = 0; i < lineCount; i++) {
                    so.selectLine({ sublistId: 'item', line: i });
                    var itemId = so.getCurrentSublistValue({ sublistId: 'item', fieldId: 'item' });
                    var qty = Number(so.getCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity' }) || 0);
                    var itemWeight = Number(itemWeights[itemId] || 0);
                    var lineWeight = itemWeight * qty;

                    totalShippingWeight += lineWeight;

                    so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'custcol_atlas_item_weight', value: itemWeight });
                    so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'custcol_atlas_line_item_weight', value: lineWeight });
                    so.commitLine({ sublistId: 'item' });
                }

                so.setValue({ fieldId: 'custbody_so_shipping_weight', value: totalShippingWeight });
                so.setValue({ fieldId: 'custbody_total_shipping_weight_kg', value: totalShippingWeight / 1000 });
            } catch (e) {
                log.error('Error populating sales order weights', e);
            }
        }

        function parseNetSuiteDate(dateStr) {
            if (!dateStr) return null;
            var t = Date.parse(dateStr);
            if (!isNaN(t)) return new Date(t);
            var parts = dateStr.split('/');
            if (parts.length === 3) {
                var m = parseInt(parts[0], 10) - 1;
                var d = parseInt(parts[1], 10);
                var y = parseInt(parts[2], 10);
                return new Date(y, m, d);
            }
            return null;
        }

        function getLastInvoicesForCustomers(customerIds) {
            var result = {};
            if (!customerIds || customerIds.length === 0) return result;
            try {
                search.create({
                    type: search.Type.INVOICE,
                    filters: [
                        ['entity', 'anyof', customerIds], 'AND',
                        ['mainline', 'is', 'T'], 'AND',
                        ['status', 'anyof', 'CustInvc:A'], 'AND', // Open / unpaid invoice
                        ['duedate', 'before', 'today']             // Past due
                    ],
                    columns: [
                        search.createColumn({ name: 'entity', summary: search.Summary.GROUP }),
                        search.createColumn({ name: 'duedate', summary: search.Summary.MIN }) // Oldest past due invoice
                    ]
                }).run().each(function (r) {
                    var custId = r.getValue({ name: 'entity', summary: search.Summary.GROUP });
                    var dateStr = r.getValue({ name: 'duedate', summary: search.Summary.MIN });
                    var days = null;
                    if (dateStr) {
                        var parsedDate = parseNetSuiteDate(dateStr);
                        if (parsedDate) {
                            var diffMs = new Date().getTime() - parsedDate.getTime();
                            days = Math.floor(diffMs / (1000 * 60 * 60 * 24));
                        }
                    }
                    result[custId] = { date: dateStr, days: days };
                    return true;
                });
            } catch (e) {
                log.error('Error in getLastInvoicesForCustomers', e);
            }
            return result;
        }

        function doCustomers(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired.' });

            var isUserAdmin = isAdmin(repId);
            if (isUserAdmin) {
                var selectedReps = req.parameters.selectedReps; // comma-separated or "all"
                var getAll = selectedReps === 'all';
                var repIds = [];
                if (!getAll && selectedReps) {
                    repIds = selectedReps.split(',').filter(Boolean);
                }
                return json(res, { ok: true, customers: getCustomersForReps(repIds, getAll) });
            } else {
                return json(res, { ok: true, customers: getCustomersForRep(repId) });
            }
        }

        /* =========================================================================
         * CRYPTO — self-contained SHA-256 + HMAC (no external setup needed)
         * ========================================================================= */
        function sha256(ascii) {
            function rr(n, x) { return (x >>> n) | (x << (32 - n)); }
            var mathPow = Math.pow, maxWord = mathPow(2, 32), result = '';
            var words = [], asciiBitLength = ascii.length * 8;
            var hash = sha256.h = sha256.h || [];
            var k = sha256.k = sha256.k || [];
            var primeCounter = k.length, isComposite = {};
            for (var candidate = 2; primeCounter < 64; candidate++) {
                if (!isComposite[candidate]) {
                    for (var i = 0; i < 313; i += candidate) isComposite[i] = candidate;
                    hash[primeCounter] = (mathPow(candidate, 0.5) * maxWord) | 0;
                    k[primeCounter++] = (mathPow(candidate, 1 / 3) * maxWord) | 0;
                }
            }
            ascii += '\x80';
            while (ascii.length % 64 - 56) ascii += '\x00';
            for (i = 0; i < ascii.length; i++) {
                var j = ascii.charCodeAt(i);
                if (j >> 8) return '';
                words[i >> 2] |= j << ((3 - i) % 4) * 8;
            }
            words[words.length] = (asciiBitLength / maxWord) | 0;
            words[words.length] = asciiBitLength;
            for (j = 0; j < words.length;) {
                var w = words.slice(j, j += 16), oldHash = hash;
                hash = hash.slice(0, 8);
                for (i = 0; i < 64; i++) {
                    var w15 = w[i - 15], w2 = w[i - 2];
                    var a = hash[0], e = hash[4];
                    var temp1 = hash[7] +
                        (rr(6, e) ^ rr(11, e) ^ rr(25, e)) +
                        ((e & hash[5]) ^ (~e & hash[6])) + k[i] +
                        (w[i] = i < 16 ? w[i] : (
                            w[i - 16] +
                            (rr(7, w15) ^ rr(18, w15) ^ (w15 >>> 3)) +
                            w[i - 7] +
                            (rr(17, w2) ^ rr(19, w2) ^ (w2 >>> 10))
                        ) | 0);
                    var temp2 = (rr(2, a) ^ rr(13, a) ^ rr(22, a)) +
                        ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));
                    hash = [(temp1 + temp2) | 0].concat(hash);
                    hash[4] = (hash[4] + temp1) | 0;
                }
                for (i = 0; i < 8; i++) hash[i] = (hash[i] + oldHash[i]) | 0;
            }
            for (i = 0; i < 8; i++) {
                for (j = 3; j + 1; j--) {
                    var b = (hash[i] >> (j * 8)) & 255;
                    result += ((b < 16) ? 0 : '') + b.toString(16);
                }
            }
            return result;
        }
        function hexToBytes(hex) { var b = []; for (var i = 0; i < hex.length; i += 2) b.push(parseInt(hex.substr(i, 2), 16)); return b; }
        function bytesToStr(b) { var s = ''; for (var i = 0; i < b.length; i++) s += String.fromCharCode(b[i]); return s; }
        function hmacSha256(key, msg) {
            var blockSize = 64;
            var keyBytes = [];
            for (var i = 0; i < key.length; i++) keyBytes.push(key.charCodeAt(i) & 0xff);
            if (keyBytes.length > blockSize) keyBytes = hexToBytes(sha256(bytesToStr(keyBytes)));
            while (keyBytes.length < blockSize) keyBytes.push(0);
            var oKey = [], iKey = [];
            for (i = 0; i < blockSize; i++) { oKey.push(keyBytes[i] ^ 0x5c); iKey.push(keyBytes[i] ^ 0x36); }
            var inner = sha256(bytesToStr(iKey) + msg);
            return sha256(bytesToStr(oKey) + bytesToStr(hexToBytes(inner)));
        }

        function doDraftOpportunity(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            var customerId = req.parameters.customerId;
            var orderType = req.parameters.orderType || '';
            var poNum = req.parameters.po || '';
            var notes = req.parameters.notes || '';
            var linesStr = req.parameters.lines || '[]';
            var selectedSalesRep = req.parameters.selectedSalesRep || '';
            var copiedFromSoId = req.parameters.copiedFromSoId || '';

            if (!customerId) return json(res, { ok: false, msg: 'No customer selected.' });

            var lines = [];
            try {
                lines = JSON.parse(linesStr);
            } catch (e) {
                return json(res, { ok: false, msg: 'Invalid lines format.' });
            }

            if (lines.length === 0) {
                return json(res, { ok: false, msg: 'No items in the order.' });
            }

            var isUserAdmin = isAdmin(repId);
            // Verify access
            if (!isUserAdmin && !repOwnsCustomer(repId, customerId)) {
                return json(res, { ok: false, msg: 'Access denied.' });
            }

            var targetRepId;
            if (isUserAdmin) {
                if (!selectedSalesRep) {
                    return json(res, { ok: false, msg: 'Please select a sales rep.' });
                }
                targetRepId = selectedSalesRep;
            } else {
                targetRepId = repId;
            }

            try {
                // Create Opportunity record
                var opp = record.create({ type: record.Type.OPPORTUNITY, isDynamic: true });
                opp.setValue({ fieldId: 'entity', value: customerId });
                opp.setValue({ fieldId: 'salesrep', value: targetRepId });

                // Determine location and address details
                var locationId = '';
                var shipAddressId = '';
                var billAddressId = '';

                if (copiedFromSoId) {
                    try {
                        var prevSo = record.load({
                            type: record.Type.SALES_ORDER,
                            id: copiedFromSoId
                        });
                        locationId = prevSo.getValue({ fieldId: 'location' }) || '';
                        shipAddressId = prevSo.getValue({ fieldId: 'shipaddresslist' }) || '';
                        billAddressId = prevSo.getValue({ fieldId: 'billaddresslist' }) || '';
                    } catch (e) {
                        log.error('Error looking up copied order details', e);
                    }
                }

                if (!locationId) {
                    locationId = resolveOrderLocationId();
                }
                if (!shipAddressId || !billAddressId) {
                    var defAddrs = getCustomerDefaultAddresses(customerId);
                    if (!shipAddressId) shipAddressId = defAddrs.shippingId;
                    if (!billAddressId) billAddressId = defAddrs.billingId;
                }

                if (locationId) {
                    try { opp.setValue({ fieldId: 'location', value: locationId }); } catch (e) { log.error('opp header location', e); }
                }
                if (shipAddressId) {
                    try { opp.setValue({ fieldId: 'shipaddresslist', value: shipAddressId }); } catch (e) { log.error('opp shipaddresslist', e); }
                }
                if (billAddressId) {
                    try { opp.setValue({ fieldId: 'billaddresslist', value: billAddressId }); } catch (e) { log.error('opp billaddresslist', e); }
                }

                if (poNum) {
                    try { opp.setValue({ fieldId: 'custbody_po_num', value: poNum }); } catch (e) { log.error('opp custbody_po_num', e); }
                    try { opp.setValue({ fieldId: 'otherrefnum', value: poNum }); } catch (e) { log.error('opp otherrefnum', e); }
                }
                if (notes) {
                    try { opp.setValue({ fieldId: 'memo', value: notes }); } catch (e) { log.error('opp memo', e); }
                }

                if (orderType) {
                    try { opp.setValue({ fieldId: 'custbody_order_type_cp', value: orderType }); } catch (e) { log.error('opp order type', e); }
                }

                opp.setValue({ fieldId: 'custbody_approve_by_admin', value: false });

                var isChristmas = isChristmasOrderType(orderType);

                lines.forEach(function (ln) {
                    var qty = Number(ln.qty || 0);
                    if (!ln.itemId) return;

                    var isOOS = !!ln.isOutOfStock || (ln.restrictionLevel === 0);
                    if (isChristmas && (qty === 0 || isOOS)) {
                        // allow it
                    } else {
                        if (!isOOS && !(qty > 0)) return;
                    }

                    opp.selectNewLine({ sublistId: 'item' });
                    opp.setCurrentSublistValue({ sublistId: 'item', fieldId: 'item', value: ln.itemId });

                    if (isChristmas && (qty === 0 || isOOS)) {
                        opp.setCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity', value: 0 });
                    } else if (isOOS) {
                        opp.setCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity', value: 0 });
                        opp.setCurrentSublistValue({ sublistId: 'item', fieldId: 'isclosed', value: true });
                    } else {
                        opp.setCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity', value: qty });
                    }

                    if (locationId) {
                        try { opp.setCurrentSublistValue({ sublistId: 'item', fieldId: 'location', value: locationId }); } catch (e) { log.error('opp line location', e); }
                    }
                    if (ln.rate != null && ln.rate !== '') {
                        opp.setCurrentSublistValue({ sublistId: 'item', fieldId: 'rate', value: Number(ln.rate) });
                    }
                    opp.commitLine({ sublistId: 'item' });
                });

                var oppId = opp.save({ enableSourcing: true, ignoreMandatoryFields: false });

                var oppTranId = search.lookupFields({
                    type: record.Type.OPPORTUNITY,
                    id: oppId,
                    columns: ['tranid']
                }).tranid;

                // Send email notification to admin users
                try {
                    sendOpportunityDraftEmail(targetRepId, customerId, oppId, oppTranId, lines, poNum, notes, orderType);
                } catch (e) {
                    log.error('Error sending opportunity draft email', e);
                }

                return json(res, { ok: true, isOpportunity: true, msg: 'Request to override sent to admin (Draft #' + oppTranId + ').', oppId: oppId, oppTranId: oppTranId });
            } catch (e) {
                log.error('Error creating Opportunity draft', e);
                return json(res, { ok: false, msg: 'Due to some issue draft is not created. Please contact administrator.' });
            }
        }

        function sendOpportunityDraftEmail(repId, customerId, oppId, oppTranId, lines, poNum, notes, orderType) {
            var conf = cfg();
            var adminUsers = conf.adminLogin;
            if (!adminUsers) {
                log.audit('sendOpportunityDraftEmail', 'No admin users configured (custscript_ord_approval_email). Email skipped.');
                return;
            }

            var recipients = adminUsers.split(',').map(function (e) { return e.trim(); });

            // Look up customer and rep names
            var custName = customerId;
            try {
                var custFields = search.lookupFields({
                    type: search.Type.CUSTOMER,
                    id: customerId,
                    columns: ['companyname', 'entityid']
                });
                custName = custFields.companyname || custFields.entityid;
            } catch (e) { }

            var repName = 'Representative';
            try {
                var repFields = search.lookupFields({
                    type: search.Type.EMPLOYEE,
                    id: repId,
                    columns: ['firstname', 'lastname', 'entityid']
                });
                repName = ((repFields.firstname || '') + ' ' + (repFields.lastname || '')).trim() || repFields.entityid;
            } catch (e) { }

            var portalUrl = url.resolveScript({
                scriptId: runtime.getCurrentScript().id,
                deploymentId: runtime.getCurrentScript().deploymentId,
                returnExternalUrl: true
            });
            var approvalUrl = portalUrl + '&oppId=' + oppId;

            var emailSubject = 'Order Portal Approval Request: Draft ' + oppTranId + ' from ' + repName;

            var isChristmas = isChristmasOrderType(orderType);

            var itemsHtml = '<table border="1" cellpadding="6" style="border-collapse:collapse; font-size:13px; font-family:sans-serif;">' +
                '<thead><tr style="background:#f1f5f9;"><th>Item</th><th>Quantity</th><th>Rate</th><th>Total</th></tr></thead><tbody>';
            var grandTotal = 0;
            lines.forEach(function (ln) {
                var qty = Number(ln.qty || 0);
                if (isChristmas) {
                    if (qty < 0) return;
                } else {
                    if (qty <= 0) return;
                }
                var rate = Number(ln.rate || 0);
                var total = qty * rate;
                grandTotal += total;
                itemsHtml += '<tr><td>' + escHtml(ln.itemName || ('Item #' + ln.itemId)) + '</td><td>' + qty + '</td><td>$' + rate.toFixed(2) + '</td><td>$' + total.toFixed(2) + '</td></tr>';
            });
            itemsHtml += '<tr style="font-weight:bold; background:#fafbfc;"><td colspan="3" style="text-align:right;">Grand Total:</td><td>$' + grandTotal.toFixed(2) + '</td></tr></tbody></table>';

            var emailBody = '<div style="font-family:sans-serif; font-size:14px; color:#0f172a;">' +
                '<p>Hello Admin,</p>' +
                '<p>A new draft order has been submitted for approval through the Order Portal because it exceeds the representative\'s order threshold limit.</p>' +
                '<hr style="border:none; border-top:1px solid #cbd5e1; margin:16px 0;" />' +
                '<p><strong>Opportunity Ref #:</strong> ' + oppTranId + '</p>' +
                '<p><strong>Sales Representative:</strong> ' + repName + '</p>' +
                '<p><strong>Customer:</strong> ' + custName + '</p>' +
                (poNum ? '<p><strong>PO Number:</strong> ' + poNum + '</p>' : '') +
                (notes ? '<p><strong>Notes/Memo:</strong> ' + notes + '</p>' : '') +
                '<p><strong>Order Items:</strong></p>' +
                itemsHtml +
                '<p style="margin-top:24px; margin-bottom:24px;">' +
                '  <a href="' + approvalUrl + '" style="background-color:#00692e; color:#ffffff; padding:12px 24px; text-decoration:none; border-radius:6px; font-weight:bold; display:inline-block;">Review &amp; Action Request</a>' +
                '</p>' +
                '<p>Please log in to the Order Portal to approve or deny this request.</p>' +
                '</div>';

            // Send email
            var author = conf.fromId || repId;
            email.send({
                author: author,
                recipients: recipients,
                subject: emailSubject,
                body: emailBody
            });
            log.audit('sendOpportunityDraftEmail', 'Sent approval email for ' + oppTranId + ' to ' + recipients.join(', '));
        }

        function escHtml(str) {
            if (!str) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        }

        function doGetOpportunities(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            var admin = isAdmin(repId);

            try {
                var filters = [];
                if (!admin) {
                    filters.push(['salesrep', 'anyof', repId]);
                }

                var opps = [];
                search.create({
                    type: search.Type.OPPORTUNITY,
                    filters: filters,
                    columns: [
                        'internalid', 'tranid', 'entity', 'salesrep', 'projectedtotal', 'custbody_approve_by_admin', 'custbody_po_num', 'memo', 'trandate', 'entitystatus',
                        search.createColumn({ name: 'companyname', join: 'customer' }),
                        search.createColumn({ name: 'firstname', join: 'salesrep' }),
                        search.createColumn({ name: 'lastname', join: 'salesrep' }),
                        search.createColumn({ name: 'entityid', join: 'salesrep' }),
                        search.createColumn({ name: 'memomain' })
                    ]
                }).run().each(function (r) {
                    var oppId = r.getValue('internalid');
                    var statusText = r.getText('entitystatus') || '';

                    // Filter out already closed/won/lost opportunities
                    if (statusText.indexOf('Closed') >= 0 || statusText.indexOf('Won') >= 0 || statusText.indexOf('Lost') >= 0) {
                        return true;
                    }
                    var notes = r.getValue({ name: 'memomain' });

                    var custName = r.getValue({ name: 'companyname', join: 'customer' }) || r.getText('entity');
                    var repName = ((r.getValue({ name: 'firstname', join: 'salesrep' }) || '') + ' ' + (r.getValue({ name: 'lastname', join: 'salesrep' }) || '')).trim() || r.getValue({ name: 'entityid', join: 'salesrep' });

                    var isApproved = r.getValue('custbody_approve_by_admin') === true || r.getValue('custbody_approve_by_admin') === 'T';

                    opps.push({
                        id: oppId,
                        tranId: r.getValue('tranid'),
                        customerId: r.getValue('entity'),
                        customerName: custName,
                        repId: r.getValue('salesrep'),
                        repName: repName,
                        notes: notes,
                        amount: Number(r.getValue('projectedtotal') || 0),
                        isApproved: isApproved,
                        po: r.getValue('custbody_po_num') || '',
                        memo: r.getValue('memo') || '',
                        date: r.getValue('trandate'),
                        orderType: ''
                    });
                    return true;
                });

                return json(res, { ok: true, opportunities: opps });
            } catch (e) {
                log.error('Error fetching opportunities', e);
                return json(res, { ok: false, msg: 'Failed to fetch draft opportunities: ' + (e.message || e) });
            }
        }

        function doApproveOpportunity(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            if (!isAdmin(repId)) {
                return json(res, { ok: false, msg: 'Access denied.' });
            }

            var oppId = req.parameters.oppId;
            if (!oppId) return json(res, { ok: false, msg: 'No opportunity selected.' });

            try {
                record.submitFields({
                    type: record.Type.OPPORTUNITY,
                    id: oppId,
                    values: {
                        custbody_approve_by_admin: true
                    }
                });
                return json(res, { ok: true, msg: 'Opportunity approved successfully.' });
            } catch (e) {
                log.error('Error approving opportunity', e);
                return json(res, { ok: false, msg: 'Failed to approve opportunity: ' + (e.message || e) });
            }
        }

        function doDenyOpportunity(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            if (!isAdmin(repId)) {
                return json(res, { ok: false, msg: 'Access denied.' });
            }

            var oppId = req.parameters.oppId;
            if (!oppId) return json(res, { ok: false, msg: 'No opportunity selected.' });

            try {
                record.delete({
                    type: record.Type.OPPORTUNITY,
                    id: oppId
                });
                return json(res, { ok: true, msg: 'Opportunity draft deleted successfully.' });
            } catch (e) {
                log.error('Error deleting opportunity', e);
                return json(res, { ok: false, msg: 'Failed to delete opportunity: ' + (e.message || e) });
            }
        }

        function doSubmitApprovedOpportunity(req, res) {
            var repId = verifyToken(req.parameters.token);
            if (!repId) return json(res, { ok: false, msg: 'Session expired. Please log in again.' });

            var oppId = req.parameters.oppId;
            if (!oppId) return json(res, { ok: false, msg: 'No opportunity selected.' });

            try {
                // Verify opportunity status and access
                var oppRec = record.load({
                    type: record.Type.OPPORTUNITY,
                    id: oppId
                });

                var isApproved = oppRec.getValue({ fieldId: 'custbody_approve_by_admin' }) === true || oppRec.getValue({ fieldId: 'custbody_approve_by_admin' }) === 'T';
                var oppRepId = oppRec.getValue({ fieldId: 'salesrep' });
                var statusText = oppRec.getText({ fieldId: 'entitystatus' }) || oppRec.getValue({ fieldId: 'entitystatus' }) || '';
                var oppLocation = oppRec.getValue({ fieldId: 'location' }) || '';
                var oppShipAddress = oppRec.getValue({ fieldId: 'shipaddresslist' }) || '';
                var oppBillAddress = oppRec.getValue({ fieldId: 'billaddresslist' }) || '';
                var oppPo = oppRec.getValue({ fieldId: 'custbody_po_num' }) || oppRec.getValue({ fieldId: 'otherrefnum' }) || '';
                var oppMemo = oppRec.getValue({ fieldId: 'memo' }) || '';
                var oppOrderType = oppRec.getValue({ fieldId: 'custbody_order_type_cp' }) || '';

                if (statusText.indexOf('Closed') >= 0 || statusText.indexOf('Won') >= 0 || statusText.indexOf('Lost') >= 0) {
                    return json(res, { ok: false, msg: 'This opportunity has already been closed or converted to a Sales Order.' });
                }

                if (!isApproved) {
                    return json(res, { ok: false, msg: 'This opportunity is not approved yet.' });
                }

                if (!isAdmin(repId) && String(repId) !== String(oppRepId)) {
                    return json(res, { ok: false, msg: 'Access denied. You can only submit your own approved drafts.' });
                }

                // Transform Opportunity to Sales Order
                var so = record.transform({
                    fromType: record.Type.OPPORTUNITY,
                    fromId: oppId,
                    toType: record.Type.SALES_ORDER,
                    isDynamic: true
                });

                if (oppLocation) {
                    try { so.setValue({ fieldId: 'location', value: oppLocation }); } catch (e) { log.error('transform header location', e); }
                }
                if (oppShipAddress) {
                    try { so.setValue({ fieldId: 'shipaddresslist', value: oppShipAddress }); } catch (e) { log.error('transform shipaddresslist', e); }
                }
                if (oppBillAddress) {
                    try { so.setValue({ fieldId: 'billaddresslist', value: oppBillAddress }); } catch (e) { log.error('transform billaddresslist', e); }
                }
                if (oppPo) {
                    try { so.setValue({ fieldId: 'otherrefnum', value: oppPo }); } catch (e) { log.error('transform header po', e); }
                }
                if (oppMemo) {
                    try { so.setValue({ fieldId: 'memo', value: oppMemo }); } catch (e) { log.error('transform header memo', e); }
                }
                if (oppOrderType) {
                    try { so.setValue({ fieldId: 'custbody_order_type_cp', value: oppOrderType }); } catch (e) { log.error('transform header order type', e); }
                }

                // Set order status to B (Pending Fulfillment) since it was approved
                so.setValue({ fieldId: 'orderstatus', value: 'B' });

                var lineCount = so.getLineCount({ sublistId: 'item' });
                for (var i = 0; i < lineCount; i++) {
                    so.selectLine({ sublistId: 'item', line: i });
                    try {
                        var qty = Number(so.getCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity' }) || 0);
                        if (qty === 0) {
                            so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'isclosed', value: true });
                        }
                        if (oppLocation) {
                            so.setCurrentSublistValue({ sublistId: 'item', fieldId: 'location', value: oppLocation });
                        }
                        so.commitLine({ sublistId: 'item' });
                    } catch (e) {
                        log.error('transform line processing error at line ' + i, e);
                    }
                }

                populateSalesOrderWeights(so);

                var savedId = so.save({ enableSourcing: true, ignoreMandatoryFields: false });

                var soNum = search.lookupFields({
                    type: record.Type.SALES_ORDER,
                    id: savedId,
                    columns: ['tranid']
                }).tranid;

                return json(res, { ok: true, salesOrderId: savedId, salesOrderNum: soNum, msg: 'Sales Order ' + soNum + ' created successfully from approved draft.' });
            } catch (e) {
                log.error('Error converting approved opportunity to sales order', e);
                return json(res, { ok: false, msg: 'Failed to submit approved draft: ' + (e.message || e) });
            }
        }

        /* =========================================================================
         * RESPONSE HELPERS
         * ========================================================================= */
        function json(res, obj) {
            res.setHeader({ name: 'Content-Type', value: 'application/json' });
            res.write(JSON.stringify(obj));
        }
        function writeJsonOrHtml(req, res, ok, msg) {
            var action = (req.parameters.action || '').toLowerCase();
            var isApi = ['requestcode', 'verifycode', 'customers', 'orderform', 'itemsearch', 'submitorder', 'draftopportunity', 'getopportunities', 'approveopportunity', 'denyopportunity', 'submitapprovedopportunity', 'getcreditrequests', 'approvecreditrequest', 'denycreditrequest', 'createcreditmemo', 'logout'].indexOf(action) >= 0;
            if (isApi) return json(res, { ok: ok, msg: msg });
            res.write('<html><body style="font-family:sans-serif;padding:40px">' +
                '<h2>' + msg + '</h2></body></html>');
        }

        return { onRequest: onRequest };
    });