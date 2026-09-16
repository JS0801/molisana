/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 *
 * Vendor Bill Approval Portal (single file, custom HTML page)
 * - Validator (all users): sees only pending bills where custbody_vendbill_validator = current user.
 *   Approve ticks custbody_vendbill_validator_check, bill stays Pending for the final approver.
 * - Final approvers (employees -5 and 12138): see ALL pending bills. Approve / Reject at any stage is final.
 * - Administrator role: sees ALL pending bills. Can act only on bills where they are the validator.
 */
define(['N/search', 'N/record', 'N/runtime', 'N/redirect', 'N/url'],
    (search, record, runtime, redirect, url) => {

        // ---- Field IDs (change here if yours differ) ----
        const FLD_VALIDATOR = 'custbody_vendbill_validator';
        const FLD_VALIDATOR_APPR = 'custbody_vendbill_validator_check';  // checkbox "Validator Approved?"
        const PARAM_FINAL_APPROVER = 'custscript_bill_final_approver';    // employee on the script
        const FINAL_APPROVERS = ['-5', '12138'];                           // final approver employee IDs
        const ADMIN_ROLE = '3';                                            // Administrator role ID

        // Approval status internal IDs
        const STATUS = { PENDING: '1', APPROVED: '2', REJECTED: '3' };

        // ==================================================================
        const onRequest = (context) => {
            const req = context.request;
            const script = runtime.getCurrentScript();
            const user = runtime.getCurrentUser();
            const userId = String(user.id);
            // Final approvers = employees in FINAL_APPROVERS, plus the script parameter if set
            const paramApprover = String(script.getParameter({ name: PARAM_FINAL_APPROVER }) || '');
            const isFinal = FINAL_APPROVERS.includes(userId) || (paramApprover !== '' && userId === paramApprover);
            const isAdmin = String(user.role) === ADMIN_ROLE;
            const seeAll = isFinal || isAdmin;

            // ---- Approve / Reject button (POST) ----
            if (req.method === 'POST') {
                const result = processBill(req.parameters.bpaction, req.parameters.billid, userId, isFinal);
                redirect.toSuitelet({
                    scriptId: script.id,
                    deploymentId: script.deploymentId,
                    parameters: { msg: result.msg, msgtype: result.type }
                });
                return;
            }

            // ---- Page (GET) ----
            const slUrl = url.resolveScript({ scriptId: script.id, deploymentId: script.deploymentId });
            const bills = getBills(userId, seeAll);
            context.response.write(renderPage({
                bills, isFinal, isAdmin, userId, slUrl,
                userName: user.name,
                msg: req.parameters.msg,
                msgType: req.parameters.msgtype
            }));
        };

        // ==================================================================
        // Data
        // ==================================================================
        const getBills = (userId, seeAll) => {
            const filters = [['mainline', 'is', 'T'], 'AND', ['approvalstatus', 'anyof', STATUS.PENDING]];
            // Everyone except final approver / admin: only bills where they are the validator
            if (!seeAll) filters.push('AND', [FLD_VALIDATOR, 'anyof', userId]);

            const bills = [];
            search.create({
                type: search.Type.VENDOR_BILL,
                filters,
                columns: [
                    search.createColumn({ name: 'trandate', sort: search.Sort.DESC }),
                    'entity', 'tranid', 'amount', FLD_VALIDATOR, FLD_VALIDATOR_APPR, 'approvalstatus'
                ]
            }).run().each((r) => {
                const v = r.getValue(FLD_VALIDATOR_APPR);
                bills.push({
                    id: r.id,
                    tranId: r.getValue('tranid') || r.id,
                    vendor: r.getText('entity') || '',
                    amount: Math.abs(Number(r.getValue('amount')) || 0),
                    date: r.getValue('trandate') || '',
                    validator: r.getText(FLD_VALIDATOR) || '',
                    validatorId: String(r.getValue(FLD_VALIDATOR) || ''),
                    validated: v === true || v === 'T',
                    status: r.getText('approvalstatus') || 'Pending Approval',
                    link: url.resolveRecord({ recordType: record.Type.VENDOR_BILL, recordId: r.id })
                });
                return true;
            });
            return bills;
        };

        // ==================================================================
        // Approve / Reject one bill
        // ==================================================================
        const processBill = (action, id, userId, isFinal) => {
            if (!id || (action !== 'approve' && action !== 'reject')) {
                return { msg: 'Nothing was changed. The request was missing the bill or the action.', type: 'warning' };
            }
            const approve = action === 'approve';
            try {
                const bill = search.lookupFields({
                    type: search.Type.VENDOR_BILL, id,
                    columns: ['tranid', FLD_VALIDATOR, FLD_VALIDATOR_APPR, 'approvalstatus']
                });
                const tranId = bill.tranid || id;
                const validatorId = bill[FLD_VALIDATOR]?.[0]?.value;
                const status = bill.approvalstatus?.[0]?.value;

                if (status !== STATUS.PENDING) {
                    return { msg: `Bill ${tranId} is no longer pending, so it was not changed.`, type: 'warning' };
                }

                let values;
                if (isFinal) {
                    // Final approver: direct master decision, validator not required
                    values = { approvalstatus: approve ? STATUS.APPROVED : STATUS.REJECTED };
                } else if (String(validatorId) === userId) {
                    if (bill[FLD_VALIDATOR_APPR] === true) {
                        return { msg: `Bill ${tranId} is already validated and waiting for the final approver.`, type: 'warning' };
                    }
                    // Validator: approve = tick checkbox (email script takes it to final approver)
                    values = approve ? { [FLD_VALIDATOR_APPR]: true } : { approvalstatus: STATUS.REJECTED };
                } else {
                    return { msg: `You are not the validator for bill ${tranId}.`, type: 'warning' };
                }

                record.submitFields({
                    type: record.Type.VENDOR_BILL, id, values,
                    options: { enableSourcing: false, ignoreMandatoryFields: true }
                });

                let msg;
                if (!approve) msg = `Bill ${tranId} rejected.`;
                else if (isFinal) msg = `Bill ${tranId} approved.`;
                else msg = `Bill ${tranId} approved and sent to the final approver.`;
                return { msg, type: 'success' };
            } catch (e) {
                log.error({ title: 'Bill ' + id, details: e });
                return { msg: `Bill ${id} could not be updated: ${e.message}`, type: 'error' };
            }
        };

        // ==================================================================
        // HTML
        // ==================================================================
        const esc = (s) => String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

        const money = (n) => n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

        const renderPage = ({ bills, isFinal, isAdmin, userId, slUrl, userName, msg, msgType }) => {
            const cntValidator = bills.filter(b => !b.validated).length;
            const cntFinal = bills.filter(b => b.validated).length;
            const defaultFilter = (isFinal || isAdmin) ? 'all' : 'validator';
            const roleLabel = isFinal ? 'Final approver' : isAdmin ? 'Administrator' : 'Validator';

            const rows = bills.map((b) => {
                const stage = b.validated ? 'final' : 'validator';
                // Final approver: any stage. Others: only their own bills not yet validated.
                const canAct = isFinal || (b.validatorId === userId && !b.validated);

                const actions = canAct
                    ? `<form method="POST" action="${esc(slUrl)}" class="act">
                         <input type="hidden" name="billid" value="${esc(b.id)}">
                         <button type="submit" name="bpaction" value="approve" class="btn approve"
                                 data-confirm="Approve bill ${esc(b.tranId)}?"
                                 onclick="return confirm(this.getAttribute('data-confirm'))">Approve</button>
                         <button type="submit" name="bpaction" value="reject" class="btn reject"
                                 data-confirm="Reject bill ${esc(b.tranId)}?"
                                 onclick="return confirm(this.getAttribute('data-confirm'))">Reject</button>
                       </form>`
                    : `<span class="waiting">${b.validated ? 'With final approver' : 'Waiting for validator'}</span>`;

                return `
                <tr data-stage="${stage}" data-search="${esc((b.tranId + ' ' + b.vendor + ' ' + b.validator).toLowerCase())}">
                  <td><a class="bill" href="${esc(b.link)}" target="_blank" rel="noopener">${esc(b.tranId)}</a></td>
                  <td>${esc(b.vendor)}</td>
                  <td class="date">${esc(b.date)}</td>
                  <td class="num">${money(b.amount)}</td>
                  <td>${esc(b.validator) || '<span class="muted">Not assigned</span>'}</td>
                  <td>${b.validated ? '<span class="tag yes">Yes</span>' : '<span class="tag no">No</span>'}</td>
                  <td><span class="tag pending">${esc(b.status)}</span></td>
                  <td class="actions">${actions}</td>
                </tr>`;
            }).join('');

            const notice = msg
                ? `<div class="notice ${esc(msgType || 'success')}" role="status">
                     <span>${esc(msg)}</span>
                     <button type="button" class="close" aria-label="Dismiss" onclick="this.parentNode.remove()">&times;</button>
                   </div>`
                : '';

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
  .search{width:320px;max-width:100%;padding:8px 12px;border:1px solid var(--rule);border-radius:6px;font:inherit;background:var(--white)}
  .shown{color:var(--muted)}

  .table-box{background:var(--white);border:1px solid var(--rule);border-radius:8px;overflow-x:auto}
  table{width:100%;border-collapse:collapse;min-width:960px}
  th{text-align:left;font-weight:500;color:var(--muted);font-size:13px;padding:10px 14px;border-bottom:1px solid var(--rule);background:#FAFBFC;white-space:nowrap}
  td{padding:11px 14px;border-bottom:1px solid #EEF1F4;vertical-align:middle}
  tbody tr:last-child td{border-bottom:0}
  tbody tr:hover td{background:#FAFBFC}
  .num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
  th.num{text-align:right}
  .date{white-space:nowrap}
  .muted{color:var(--muted)}
  a.bill{color:var(--focus);font-weight:500;text-decoration:none}
  a.bill:hover{text-decoration:underline}

  .tag{display:inline-block;padding:1px 8px;border-radius:4px;font-size:12px;font-weight:500;white-space:nowrap}
  .tag.yes{background:var(--approve-bg);color:var(--approve)}
  .tag.no{background:#EEF1F4;color:var(--muted)}
  .tag.pending{background:var(--pending-bg);color:var(--pending)}

  .actions{white-space:nowrap}
  .act{display:flex;gap:6px;margin:0}
  .btn{border:1px solid transparent;border-radius:6px;padding:5px 14px;font:inherit;font-weight:500;cursor:pointer}
  .btn.approve{background:var(--approve);color:#fff}
  .btn.approve:hover{background:#18643F}
  .btn.reject{background:var(--white);color:var(--reject);border-color:#E3B3AF}
  .btn.reject:hover{background:var(--reject-bg)}
  .waiting{color:var(--muted);font-size:13px}

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
    <input type="search" class="search" id="q" placeholder="Search bill, vendor or validator" aria-label="Search bills">
    <span class="shown" id="shown"></span>
  </div>

  <div class="table-box">
    <table>
      <thead>
        <tr>
          <th>Bill</th><th>Vendor</th><th>Date</th><th class="num">Amount</th>
          <th>Validator</th><th>Validator approved?</th><th>Approval status</th><th>Action</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
    <div class="empty" id="empty" hidden>No bills match this view.</div>
  </div>

</div>

<script>
(function () {
  var rows = Array.prototype.slice.call(document.querySelectorAll('tbody tr'));
  var tabs = Array.prototype.slice.call(document.querySelectorAll('.count'));
  var q = document.getElementById('q');
  var empty = document.getElementById('empty');
  var shown = document.getElementById('shown');
  var filter = '${defaultFilter}';

  function apply() {
    var term = q.value.trim().toLowerCase(), n = 0;
    rows.forEach(function (r) {
      var ok = (filter === 'all' || r.getAttribute('data-stage') === filter) &&
               (!term || r.getAttribute('data-search').indexOf(term) > -1);
      r.hidden = !ok;
      if (ok) n++;
    });
    tabs.forEach(function (t) { t.setAttribute('aria-pressed', t.getAttribute('data-filter') === filter); });
    empty.hidden = n > 0;
    shown.textContent = n + (n === 1 ? ' bill' : ' bills') + ' shown';
  }

  tabs.forEach(function (t) {
    t.addEventListener('click', function () { filter = t.getAttribute('data-filter'); apply(); });
  });
  q.addEventListener('input', apply);
  apply();
})();
</script>
</body>
</html>`;
        };

        return { onRequest };
    });