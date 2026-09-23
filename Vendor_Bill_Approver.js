/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 *
 * Vendor Bill Approval Portal (single file, custom HTML page)
 * - Opened ONLY from the Molisana employee portal (signed portal token, same as the other tools). Direct access -> redirected to the portal login.
 * - Identity comes from the token (empid), not from a NetSuite login.
 * - Validator (all users): sees only pending bills where custbody_vendbill_validator = current user.
 *   Approve ticks custbody_vendbill_validator_check, bill stays Pending for the final approver.
 * - Final approvers (employees -5, 8, 12138): see ALL pending bills. Approve / Reject at any stage is final.
 * - View only (employees 12412, 11018, 11428): see ALL pending bills, no Approve / Reject.
 * - Bills are selected with checkboxes and approved / rejected in bulk from the buttons above the list.
 */
define(['N/search', 'N/record', 'N/runtime', 'N/format', 'N/crypto'],
    (search, record, runtime, format, crypto) => {

        // ---- Field IDs (change here if yours differ) ----
        const FLD_VALIDATOR = 'custbody_vendbill_validator';
        const FLD_VALIDATOR_APPR = 'custbody_vendbill_validator_check';   // checkbox "Validator Approved?"
        const FLD_NOTE = 'custbody_note_to_vendor';                        // "Note to Vendor"
        const FLD_DOC_URL = 'custbody_mi_sharepoint_document_url';         // SharePoint document link
        const PARAM_FINAL_APPROVER = 'custscript_bill_final_approver';     // employee on the script

        // ---- Access ----
        const FINAL_APPROVERS = ['-5', '8', '12138'];                      // full rights: approve / reject any bill
        const VIEW_ONLY = ['12412', '11018', '11428'];                     // see every bill, cannot approve / reject

        // ---- Portal gate ----
        const PARAM_SECRET = 'custscript_portal_secret';                   // same value as on the portal script
        const TOKEN_TTL_MS = 30 * 60 * 1000;                               // 30 minutes, same as the portal
        const LOGIN_URL = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2110&deploy=1&compid=4975346&ns-at=AAEJ7tMQamzukv1WMqTK6i2c27bRetbrd2MDLjhDgPPFOawMxCo';
        const SELF_URL = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=3597&deploy=1&compid=4975346&ns-at=AAEJ7tMQI72usOcKnZXjZTEWK6RnSYmy-rFCR124tEQ5npv7o2k';
        const NS_BILL_BASE = 'https://4975346.app.netsuite.com/app/accounting/transactions/vendbill.nl?id=';

        // Line description column. The first one that works in this account is used.
        const LINE_DESC_COLUMNS = ['description', 'memo'];

        // Approval status internal IDs
        const STATUS = { PENDING: '1', APPROVED: '2', REJECTED: '3' };

        // ==================================================================
        const onRequest = (context) => {
            const req = context.request;
            const script = runtime.getCurrentScript();

            // ---- Gate: valid signed token from the portal, else back to the login page ----
            const empid = String(req.parameters.empid || '');
            const ts = String(req.parameters.ts || '');
            const sig = String(req.parameters.sig || '');
            const secret = String(script.getParameter({ name: PARAM_SECRET }) || '');
            if (!verifyToken(secret, empid, ts, sig)) {
                denied(context, req.method === 'POST' ? 'Session expired. Please log in again.' : 'Login Required');
                return;
            }

            const userId = empid;
            const paramApprover = String(script.getParameter({ name: PARAM_FINAL_APPROVER }) || '');
            const isFinal = FINAL_APPROVERS.includes(userId) || (paramApprover !== '' && userId === paramApprover);
            const isViewOnly = VIEW_ONLY.includes(userId) && !isFinal;
            const seeAll = isFinal || isViewOnly;

            // Token travels on every link / form so the page keeps working (and stays gated)
            const slUrl = SELF_URL + '&empid=' + encodeURIComponent(empid) + '&ts=' + encodeURIComponent(ts) + '&sig=' + encodeURIComponent(sig);

            // ---- Bulk Approve / Reject (POST) ----
            if (req.method === 'POST') {
                const result = processBills(req.parameters.bpaction, req.parameters.billids, userId, isFinal, isViewOnly);
                redirectTo(context, slUrl
                    + '&msg=' + encodeURIComponent(result.msg) + '&msgtype=' + encodeURIComponent(result.type)
                    + '&from=' + encodeURIComponent(req.parameters.from || '') + '&to=' + encodeURIComponent(req.parameters.to || ''));
                return;
            }

            // ---- Page (GET) ----
            const from = req.parameters.from || '';   // yyyy-mm-dd from the date pickers
            const to = req.parameters.to || '';
            const bills = getBills(userId, seeAll, from, to);
            context.response.write(renderPage({
                bills, isFinal, isViewOnly, userId, slUrl, from, to,
                userName: getUserName(userId),
                msg: req.parameters.msg,
                msgType: req.parameters.msgtype
            }));
        };

        // ==================================================================
        // Portal gate helpers
        // ==================================================================
        const signToken = (secret, empid, ts) => {
            const h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
            h.update({ input: empid + '|' + ts + '|' + secret });
            return h.digest({ outputEncoding: crypto.Encoding.HEX });
        };

        const verifyToken = (secret, empid, ts, sig) => {
            if (!secret || !empid || !ts || !sig) return false;      // no secret configured = deny
            const age = Math.abs(Date.now() - parseInt(ts, 10));
            if (!(age <= TOKEN_TTL_MS)) return false;                // also catches NaN
            try { return signToken(secret, empid, ts) === sig; }
            catch (e) { log.error({ title: 'verifyToken', details: e }); return false; }
        };

        // Same "Login Required" bounce as the other portal tools
        const denied = (context, message) => {
            context.response.write(
                '<html><head>' +
                '<script>setTimeout(function(){ window.location.href = ' + JSON.stringify(LOGIN_URL) + '; }, 1200);</script>' +
                '<style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>' +
                '</head><body><div class="message">' + esc(message) + '</div></body></html>'
            );
        };

        const redirectTo = (context, target) => {
            context.response.write('<html><body><script>window.location.href=' + JSON.stringify(target) + ';</script></body></html>');
        };

        const getUserName = (id) => {
            try {
                return search.lookupFields({ type: search.Type.EMPLOYEE, id, columns: ['entityid'] }).entityid || '';
            } catch (e) { return ''; }
        };

        // ==================================================================
        // Data
        // ==================================================================

        // yyyy-mm-dd (from the date picker) -> the account's date format, for search filters
        const toNsDate = (iso) => {
            if (!iso) return '';
            const p = iso.split('-');
            if (p.length !== 3) return '';
            return format.format({ value: new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])), type: format.Type.DATE });
        };

        const getBills = (userId, seeAll, from, to) => {
            const filters = [['mainline', 'is', 'T'], 'AND', ['approvalstatus', 'anyof', STATUS.PENDING]];
            // Everyone except final approver / view only / admin: only bills where they are the validator
            if (!seeAll) filters.push('AND', [FLD_VALIDATOR, 'anyof', userId]);

            // Bill date filter
            const f = toNsDate(from), t = toNsDate(to);
            if (f && t) filters.push('AND', ['trandate', 'within', f, t]);
            else if (f) filters.push('AND', ['trandate', 'onorafter', f]);
            else if (t) filters.push('AND', ['trandate', 'onorbefore', t]);

            const bills = [];
            search.create({
                type: search.Type.VENDOR_BILL,
                filters,
                columns: [
                    search.createColumn({ name: 'trandate', sort: search.Sort.DESC }),
                    'entity', 'tranid', 'amount', FLD_NOTE, FLD_DOC_URL,
                    FLD_VALIDATOR, FLD_VALIDATOR_APPR, 'approvalstatus'
                ]
            }).run().each((r) => {
                const v = r.getValue(FLD_VALIDATOR_APPR);
                bills.push({
                    id: r.id,
                    tranId: r.getValue('tranid') || r.id,
                    vendor: r.getText('entity') || '',
                    amount: Math.abs(Number(r.getValue('amount')) || 0),
                    date: r.getValue('trandate') || '',
                    note: r.getValue(FLD_NOTE) || '',
                    docUrl: pickUrl(r.getValue(FLD_DOC_URL)),
                    validator: r.getText(FLD_VALIDATOR) || '',
                    validatorId: String(r.getValue(FLD_VALIDATOR) || ''),
                    validated: v === true || v === 'T',
                    status: r.getText('approvalstatus') || 'Pending Approval',
                    link: NS_BILL_BASE + r.id,
                    lineDesc: ''
                });
                return true;
            });

            addLineDescriptions(bills);
            return bills;
        };

        // Item line descriptions, shown in the bill's own row
        const addLineDescriptions = (bills) => {
            if (!bills.length) return;

            const byId = {};
            bills.forEach(b => { byId[b.id] = b; });

            const filters = [
                ['internalid', 'anyof', Object.keys(byId)], 'AND',
                ['mainline', 'is', 'F'], 'AND',
                ['taxline', 'is', 'F'], 'AND',
                ['shipping', 'is', 'F']
            ];

            const load = (descCol) => {
                const columns = ['internalid', 'item'];
                if (descCol) columns.push(descCol);
                const parts = {};

                search.create({ type: search.Type.VENDOR_BILL, filters, columns }).run().each((r) => {
                    const id = r.getValue({ name: 'internalid' });
                    if (!byId[id]) return true;
                    const text = (descCol ? r.getValue(descCol) : '') || r.getText('item') || '';
                    if (!text) return true;
                    if (!parts[id]) parts[id] = [];
                    if (parts[id].indexOf(text) === -1) parts[id].push(text);
                    return true;
                });

                Object.keys(parts).forEach(id => { byId[id].lineDesc = parts[id].join(' | '); });
            };

            // Try each description column, then fall back to item names only
            const tries = LINE_DESC_COLUMNS.concat([null]);
            for (let i = 0; i < tries.length; i++) {
                try {
                    load(tries[i]);
                    return;
                } catch (e) {
                    log.audit({ title: 'Line description column not usable: ' + tries[i], details: e.message });
                }
            }
        };

        // ==================================================================
        // Bulk Approve / Reject
        // ==================================================================
        const processBills = (action, billids, userId, isFinal, isViewOnly) => {
            if (isViewOnly) return { msg: 'You have view only access to this portal.', type: 'warning' };
            if (action !== 'approve' && action !== 'reject') {
                return { msg: 'Nothing was changed. No action was received.', type: 'warning' };
            }

            const ids = String(billids || '').split(',').map(s => s.trim()).filter(Boolean);
            if (!ids.length) return { msg: 'Please tick at least one bill.', type: 'warning' };

            const approve = action === 'approve';
            let done = 0, skipped = 0, errors = 0;

            ids.forEach((id) => {
                try {
                    const bill = search.lookupFields({
                        type: search.Type.VENDOR_BILL, id,
                        columns: [FLD_VALIDATOR, FLD_VALIDATOR_APPR, 'approvalstatus']
                    });
                    const validatorId = bill[FLD_VALIDATOR]?.[0]?.value;
                    const status = bill.approvalstatus?.[0]?.value;

                    if (status !== STATUS.PENDING) { skipped++; return; }

                    let values;
                    if (isFinal) {
                        // Final approver: direct master decision, validator not required
                        values = { approvalstatus: approve ? STATUS.APPROVED : STATUS.REJECTED };
                    } else if (String(validatorId) === userId && bill[FLD_VALIDATOR_APPR] !== true) {
                        // Validator: approve = tick checkbox (email script takes it to the final approver)
                        values = approve ? { [FLD_VALIDATOR_APPR]: true } : { approvalstatus: STATUS.REJECTED };
                    } else {
                        skipped++;
                        return;
                    }

                    record.submitFields({
                        type: record.Type.VENDOR_BILL, id, values,
                        options: { enableSourcing: false, ignoreMandatoryFields: true }
                    });
                    done++;
                } catch (e) {
                    errors++;
                    log.error({ title: 'Bill ' + id, details: e });
                }
            });

            const what = done === 1 ? 'bill' : 'bills';
            let msg = `${approve ? 'Approved' : 'Rejected'} ${done} ${what}.`;
            if (approve && !isFinal && done) msg += ' Sent to the final approver.';
            if (skipped) msg += ` ${skipped} skipped (no longer pending or not yours).`;
            if (errors) msg += ` ${errors} failed, see the script log.`;

            return { msg, type: errors ? 'error' : (done ? 'success' : 'warning') };
        };

        // ==================================================================
        // HTML
        // ==================================================================
        const esc = (s) => String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

        // The field may hold a plain URL or a hyperlink; keep only an http(s) address
        const pickUrl = (v) => {
            const m = String(v == null ? '' : v).match(/https?:\/\/[^\s"'<>]+/);
            return m ? m[0] : '';
        };

        const money = (n) => n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

        const renderPage = ({ bills, isFinal, isViewOnly, userId, slUrl, from, to, userName, msg, msgType }) => {
            const cntValidator = bills.filter(b => !b.validated).length;
            const cntFinal = bills.filter(b => b.validated).length;
            const defaultFilter = (isFinal || isViewOnly) ? 'all' : 'validator';
            const roleLabel = isFinal ? 'Final approver' : isViewOnly ? 'View only' : 'Validator';

            const rows = bills.map((b) => {
                const stage = b.validated ? 'final' : 'validator';
                const canAct = !isViewOnly && (isFinal || (b.validatorId === userId && !b.validated));
                const why = isViewOnly ? 'View only access'
                    : b.validated ? 'Waiting for the final approver'
                    : 'You are not the validator';

                const pick = canAct
                    ? `<input type="checkbox" class="pick" value="${esc(b.id)}" aria-label="Select bill ${esc(b.tranId)}">`
                    : `<span class="muted lock" title="${esc(why)}">&mdash;</span>`;

                return `
                <tr data-stage="${stage}" data-search="${esc((b.tranId + ' ' + b.vendor + ' ' + b.validator + ' ' + b.note + ' ' + b.lineDesc).toLowerCase())}">
                  <td class="pickcell">${pick}</td>
                  <td><a class="bill" href="${esc(b.link)}" target="_blank" rel="noopener">${esc(b.tranId)}</a></td>
                  <td>${esc(b.vendor)}</td>
                  <td class="wide" title="${esc(b.note)}"><span class="clamp">${esc(b.note) || '<span class="muted">&mdash;</span>'}</span></td>
                  <td class="wide" title="${esc(b.lineDesc)}"><span class="clamp">${esc(b.lineDesc) || '<span class="muted">&mdash;</span>'}</span></td>
                  <td>${b.docUrl
                        ? `<a class="doc" href="${esc(b.docUrl)}" target="_blank" rel="noopener">Open</a>`
                        : '<span class="muted">&mdash;</span>'}</td>
                  <td class="date">${esc(b.date)}</td>
                  <td class="num">${money(b.amount)}</td>
                  <td class="nowrap">${esc(b.validator) || '<span class="muted">Not assigned</span>'}</td>
                  <td>${b.validated ? '<span class="tag yes">Yes</span>' : '<span class="tag no">No</span>'}</td>
                  <td><span class="tag pending">${esc(b.status)}</span></td>
                </tr>`;
            }).join('');

            const notice = msg
                ? `<div class="notice ${esc(msgType || 'success')}" role="status">
                     <span>${esc(msg)}</span>
                     <button type="button" class="close" aria-label="Dismiss" onclick="this.parentNode.remove()">&times;</button>
                   </div>`
                : '';

            const bulkBar = isViewOnly ? '' : `
      <div class="bulk">
        <span class="picked" id="picked">No bills selected</span>
        <button type="button" class="btn approve" id="bulkApprove" disabled>Approve selected</button>
        <button type="button" class="btn reject" id="bulkReject" disabled>Reject selected</button>
      </div>`;

            return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Bill approvals</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  :root{
    --ink:#1B2A41; --text:#2F3B4C; --muted:#6B7686; --rule:#DDE2E8; --paper:#F5F7F9; --white:#FFFFFF;
    --approve:#1F7A4D; --approve-bg:#E6F2EC; --reject:#B3261E; --reject-bg:#FBEAE8;
    --pending:#9A6412; --pending-bg:#FBF1DF; --focus:#2F6FDB;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--paper);color:var(--text);
       font:14px/1.5 "IBM Plex Sans",-apple-system,"Segoe UI",Roboto,Arial,sans-serif;}
  .wrap{width:100%;min-height:100vh;padding:28px 32px 48px}

  header{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap;margin-bottom:22px}
  h1{margin:0;color:var(--ink);font-size:26px;font-weight:600;letter-spacing:-.01em}
  .who{color:var(--muted);margin-top:2px}
  .role{display:inline-block;padding:2px 10px;border-radius:999px;background:var(--ink);color:#fff;font-size:12px;font-weight:500}

  /* Count strip: doubles as the filter */
  .counts{display:flex;background:var(--white);border:1px solid var(--rule);border-radius:8px;overflow:hidden;margin-bottom:16px}
  .count{flex:1;min-width:0;border:0;background:none;text-align:left;padding:14px 20px 12px;cursor:pointer;
         border-bottom:3px solid transparent;font:inherit;color:inherit}
  .count + .count{border-left:1px solid var(--rule)}
  .count .n{display:block;font-size:30px;font-weight:600;color:var(--ink);line-height:1.1;font-variant-numeric:tabular-nums}
  .count .l{display:flex;align-items:center;gap:8px;color:var(--muted)}
  .count .dot{width:8px;height:8px;border-radius:50%}
  .count:hover{background:#FAFBFC}
  .count[aria-pressed="true"]{border-bottom-color:var(--ink)}
  .count[aria-pressed="true"] .l{color:var(--ink)}

  .toolbar{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:10px;flex-wrap:wrap}
  .tools{display:flex;align-items:center;gap:18px;flex-wrap:wrap}
  .search{width:260px;max-width:100%;padding:8px 12px;border:1px solid var(--rule);border-radius:6px;font:inherit;background:var(--white)}
  .dates{display:flex;align-items:center;gap:8px;margin:0;flex-wrap:wrap}
  .dates label{color:var(--muted)}
  .dates input[type=date]{padding:7px 10px;border:1px solid var(--rule);border-radius:6px;font:inherit;background:var(--white);color:inherit}
  .dates .to{color:var(--muted)}
  .clear{color:var(--muted);text-decoration:none;font-size:13px}
  .clear:hover{text-decoration:underline}

  .bulk{display:flex;align-items:center;gap:10px}
  .picked{color:var(--muted)}
  .btn{border:1px solid transparent;border-radius:6px;padding:7px 16px;font:inherit;font-weight:500;cursor:pointer}
  .btn.plain{background:var(--white);border-color:var(--rule);color:var(--ink)}
  .btn.plain:hover{background:#F0F3F6}
  .btn.approve{background:var(--approve);color:#fff}
  .btn.approve:hover{background:#18643F}
  .btn.reject{background:var(--white);color:var(--reject);border-color:#E3B3AF}
  .btn.reject:hover{background:var(--reject-bg)}
  .btn:disabled{opacity:.45;cursor:default}
  .btn.approve:disabled:hover{background:var(--approve)}
  .btn.reject:disabled:hover{background:var(--white)}

  .table-box{background:var(--white);border:1px solid var(--rule);border-radius:8px;overflow-x:auto}
  table{width:100%;border-collapse:collapse;min-width:1240px}
  th{text-align:left;font-weight:500;color:var(--muted);font-size:13px;padding:10px 14px;border-bottom:1px solid var(--rule);background:#FAFBFC;white-space:nowrap}
  td{padding:11px 14px;border-bottom:1px solid #EEF1F4;vertical-align:middle}
  tbody tr:last-child td{border-bottom:0}
  tbody tr:hover td{background:#FAFBFC}
  .num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
  th.num{text-align:right}
  .date{white-space:nowrap}
  .muted{color:var(--muted)}
  .lock{cursor:help}
  .pickcell{width:40px;text-align:center}
  input[type=checkbox]{width:16px;height:16px;cursor:pointer}
  td.wide{max-width:280px}
  .clamp{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}

  .nowrap{white-space:nowrap}
  a.bill{color:var(--focus);font-weight:500;text-decoration:none;white-space:nowrap}
  a.bill:hover{text-decoration:underline}
  a.doc{display:inline-block;padding:2px 10px;border:1px solid var(--rule);border-radius:4px;
        color:var(--focus);text-decoration:none;font-size:13px;white-space:nowrap}
  a.doc:hover{background:#F0F3F6}

  .tag{display:inline-block;padding:1px 8px;border-radius:4px;font-size:12px;font-weight:500;white-space:nowrap}
  .tag.yes{background:var(--approve-bg);color:var(--approve)}
  .tag.no{background:#EEF1F4;color:var(--muted)}
  .tag.pending{background:var(--pending-bg);color:var(--pending)}

  .empty{padding:48px 20px;text-align:center;color:var(--muted)}

  .notice{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:10px 14px;border-radius:6px;margin-bottom:16px;border:1px solid}
  .notice.success{background:var(--approve-bg);border-color:#BFDCCB;color:#175C3A}
  .notice.warning{background:var(--pending-bg);border-color:#EED8B0;color:#7A4F0E}
  .notice.error{background:var(--reject-bg);border-color:#EFC5C1;color:#8C1D17}
  .close{border:0;background:none;font-size:20px;line-height:1;color:inherit;cursor:pointer}

  button:focus-visible,a:focus-visible,input:focus-visible{outline:2px solid var(--focus);outline-offset:2px}

  @media (max-width:720px){
    .counts{flex-direction:column}
    .count + .count{border-left:0;border-top:1px solid var(--rule)}
    .search{width:100%}
  }
</style>
</head>
<body>
<div class="wrap">

  <header>
    <div>
      <h1>Bill approvals</h1>
      <div class="who">${esc(userName)}</div>
    </div>
    <span class="role">${roleLabel}</span>
  </header>

  ${notice}

  <div class="counts" role="group" aria-label="Filter bills">
    <button type="button" class="count" data-filter="all">
      <span class="n">${bills.length}</span>
      <span class="l"><span class="dot" style="background:var(--ink)"></span>Pending bills</span>
    </button>
    <button type="button" class="count" data-filter="validator">
      <span class="n">${cntValidator}</span>
      <span class="l"><span class="dot" style="background:var(--pending)"></span>Pending validator approval</span>
    </button>
    <button type="button" class="count" data-filter="final">
      <span class="n">${cntFinal}</span>
      <span class="l"><span class="dot" style="background:var(--approve)"></span>Pending final approval</span>
    </button>
  </div>

  <div class="toolbar">
    <div class="tools">
      <input type="search" class="search" id="q" placeholder="Search bill, vendor, note or description" aria-label="Search bills">
      <div class="dates">
        <label for="from">Bill date</label>
        <input type="date" id="from" value="${esc(from)}" aria-label="From date">
        <span class="to">to</span>
        <input type="date" id="to" value="${esc(to)}" aria-label="To date">
        <button type="button" class="btn plain" id="applyDates">Apply</button>
        ${(from || to) ? `<a class="clear" href="${esc(slUrl)}">Clear</a>` : ''}
      </div>
    </div>
    ${bulkBar}
  </div>

  <form method="POST" action="${esc(slUrl)}" id="bulkForm">
    <input type="hidden" name="bpaction" id="bpaction" value="">
    <input type="hidden" name="billids" id="billids" value="">
    <input type="hidden" name="from" value="${esc(from)}">
    <input type="hidden" name="to" value="${esc(to)}">

    <div class="table-box">
      <table>
        <thead>
          <tr>
            <th class="pickcell"><input type="checkbox" id="pickAll" aria-label="Select all bills"></th>
            <th>Bill</th><th>Vendor</th><th>Note to vendor</th><th>Item description</th>
            <th>SharePoint document</th><th>Date</th><th class="num">Amount</th>
            <th>Validator</th><th>Validator approved?</th><th>Approval status</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="empty" id="empty" hidden>No bills match this view.</div>
    </div>
  </form>

</div>

<script>
(function () {
  var rows = Array.prototype.slice.call(document.querySelectorAll('tbody tr'));
  var tabs = Array.prototype.slice.call(document.querySelectorAll('.count'));
  var q = document.getElementById('q');
  var empty = document.getElementById('empty');
  var pickAll = document.getElementById('pickAll');
  var picked = document.getElementById('picked');
  var filter = '${defaultFilter}';

  function boxes(visibleOnly) {
    return rows.filter(function (r) { return !visibleOnly || !r.hidden; })
               .map(function (r) { return r.querySelector('.pick'); })
               .filter(Boolean);
  }

  function chosen() {
    return boxes(true).filter(function (c) { return c.checked; }).map(function (c) { return c.value; });
  }

  function refresh() {
    if (!picked) return;
    var n = chosen().length;
    picked.textContent = n ? n + (n === 1 ? ' bill selected' : ' bills selected') : 'No bills selected';
    document.getElementById('bulkApprove').disabled = !n;
    document.getElementById('bulkReject').disabled = !n;
    var all = boxes(true);
    pickAll.checked = all.length > 0 && n === all.length;
  }

  function apply() {
    var term = q.value.trim().toLowerCase(), n = 0;
    rows.forEach(function (r) {
      var ok = (filter === 'all' || r.getAttribute('data-stage') === filter) &&
               (!term || r.getAttribute('data-search').indexOf(term) > -1);
      r.hidden = !ok;
      if (!ok) { var c = r.querySelector('.pick'); if (c) c.checked = false; }
      if (ok) n++;
    });
    tabs.forEach(function (t) { t.setAttribute('aria-pressed', t.getAttribute('data-filter') === filter); });
    empty.hidden = n > 0;
    refresh();
  }

  tabs.forEach(function (t) {
    t.addEventListener('click', function () { filter = t.getAttribute('data-filter'); apply(); });
  });
  q.addEventListener('input', apply);

  if (pickAll) {
    pickAll.addEventListener('change', function () {
      boxes(true).forEach(function (c) { c.checked = pickAll.checked; });
      refresh();
    });
  }
  rows.forEach(function (r) {
    var c = r.querySelector('.pick');
    if (c) c.addEventListener('change', refresh);
  });

  // Bulk approve / reject
  function submit(action) {
    var ids = chosen();
    if (!ids.length) return;
    var word = action === 'approve' ? 'Approve' : 'Reject';
    if (!confirm(word + ' ' + ids.length + (ids.length === 1 ? ' bill?' : ' bills?'))) return;
    document.getElementById('bpaction').value = action;
    document.getElementById('billids').value = ids.join(',');
    document.getElementById('bulkForm').submit();
  }
  if (picked) {
    document.getElementById('bulkApprove').addEventListener('click', function () { submit('approve'); });
    document.getElementById('bulkReject').addEventListener('click', function () { submit('reject'); });
  }

  // Date filter: reload the Suitelet with from / to, keeping the script and deploy params
  var SL_URL = '${slUrl}';
  var fromEl = document.getElementById('from');
  var toEl = document.getElementById('to');

  function applyDates() {
    var u = SL_URL;
    if (fromEl.value) u += '&from=' + encodeURIComponent(fromEl.value);
    if (toEl.value) u += '&to=' + encodeURIComponent(toEl.value);
    window.location.href = u;
  }
  document.getElementById('applyDates').addEventListener('click', applyDates);
  [fromEl, toEl].forEach(function (el) {
    el.addEventListener('keydown', function (e) { if (e.key === 'Enter') applyDates(); });
  });

  apply();
})();
</script>
</body>
</html>`;
        };

        return { onRequest };
    });