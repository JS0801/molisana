/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * Molisana vendor PO response. See accompanying setup instructions.
 */
define(['N/record', 'N/format', 'N/log'], (record, format, log) => {
    const CONFIG = {
        // Verify these two IDs in your account before deployment.
        decisionField: 'custbody_mi_vendor_decision',
        notesField: 'custbody_mi_vendor_notes',
        tokenField: 'custbody_mi_vendor_link_token',
        acceptedValue: '1',
        rejectedValue: '2',
        maxNotes: 1000,
        logoUrl: 'https://4975346-sb1.app.netsuite.com/core/media/media.nl?id=4770&c=4975346_SB1&h=sGvHCgrcrHMJzoZjmKF-9Og7Y_nEEncIbLigVYoLtPhS1CXd'
    };
    const esc = value => String(value == null ? '' : value).replace(/[&<>"']/g,
        ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
    const value = (po, fieldId) => po.getValue({fieldId});
    const text = (po, fieldId) => {
        try { return po.getText({fieldId}) || value(po, fieldId) || ''; }
        catch (_) { return value(po, fieldId) || ''; }
    };
    function date(v) {
        return v instanceof Date ? format.format({value: v, type: format.Type.DATE}) : (v || '—');
    }
    function money(v) {
        const n = Number(v);
        return Number.isFinite(n) ? n.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '—';
    }
    function fail(message) { const e = new Error(message); e.publicMessage = message; throw e; }

    function page(body) {
        return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="referrer" content="no-referrer"><title>Molisana | Purchase Order Response</title><style>
        :root{--navy:#172749;--gold:#b29344;--ink:#22304a;--muted:#697586;--line:#e4e7ec}*{box-sizing:border-box}body{margin:0;background:#f4f3ef;color:var(--ink);font:15px/1.6 Arial,Helvetica,sans-serif}header{background:white;border-top:5px solid var(--navy);border-bottom:1px solid var(--line)}.brand{max-width:1080px;margin:auto;padding:20px 28px;display:flex;align-items:center;justify-content:space-between;gap:20px}.brand img{width:180px;height:auto}.brand-label{text-align:right;color:var(--navy);font-size:12px;letter-spacing:2px;text-transform:uppercase}.brand-label strong{display:block;letter-spacing:0;font-size:16px;margin-bottom:3px}main{max-width:1080px;margin:36px auto;padding:0 24px}.eyebrow{font-size:11px;font-weight:bold;letter-spacing:2px;color:#806624;text-transform:uppercase}h1{font-size:34px;line-height:1.2;margin:8px 0 12px;color:var(--navy)}h2{font-size:18px;margin:0 0 18px}.sub{color:var(--muted);margin:0 0 26px}.card{background:white;border:1px solid var(--line);border-radius:14px;box-shadow:0 5px 20px #14234105;margin-bottom:22px;overflow:hidden}.pad{padding:26px}.topline{display:flex;align-items:center;justify-content:space-between;gap:16px}.badge{border-radius:30px;padding:7px 14px;font-size:12px;font-weight:bold;background:#f4f0e4;color:#7a6226}.accept{background:#e8f4ed;color:#216744}.reject{background:#fbece9;color:#a3362b}.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:24px;margin-top:24px}.label{display:block;font-size:11px;letter-spacing:1px;color:var(--muted);text-transform:uppercase;margin-bottom:5px}.val{font-weight:600;overflow-wrap:anywhere}.address{white-space:pre-line;font-weight:normal}.scroll{overflow:auto}table{width:100%;border-collapse:collapse;min-width:640px;font-size:13px}th{background:#f8f9fb;color:var(--muted);font-size:10px;letter-spacing:1px;text-transform:uppercase;text-align:left;padding:13px 20px}td{padding:16px 20px;border-top:1px solid #eef0f3;vertical-align:top;overflow-wrap:anywhere}.num{text-align:right;white-space:nowrap}.description{color:var(--muted);font-size:12px;white-space:pre-line;max-width:480px}.total{display:flex;justify-content:flex-end;align-items:center;gap:35px;padding:20px 26px;background:#faf9f5;border-top:1px solid var(--line)}.total strong{font-size:24px;color:var(--navy)}label{display:block;font-weight:bold;margin:18px 0 8px}label span{font-weight:normal;color:var(--muted);font-size:13px}textarea{width:100%;min-height:135px;resize:vertical;padding:14px;border:1px solid #cdd3dd;border-radius:8px;font:inherit;color:var(--ink)}textarea:focus{outline:2px solid var(--gold);outline-offset:2px}.hint{font-size:12px;color:var(--muted);margin:6px 0 0}.actions{display:flex;justify-content:space-between;align-items:center;gap:20px;margin-top:24px}.actions p{color:var(--muted);font-size:12px;max-width:510px;margin:0}button{border:0;border-radius:8px;background:var(--navy);color:white;padding:15px 24px;font:600 14px Arial;cursor:pointer;white-space:nowrap}button.danger{background:#a63c31}button:hover{filter:brightness(1.12)}.notice{padding:18px 20px;border-radius:8px;background:#f5f1e7;border-left:3px solid var(--gold);margin-bottom:20px}.success{padding:46px;text-align:center}.icon{display:inline-flex;align-items:center;justify-content:center;border-radius:50%;width:58px;height:58px;font-size:27px;margin-bottom:18px;background:#e8f4ed;color:#216744}footer{text-align:center;padding:10px 24px 32px;color:var(--muted);font-size:11px}a{color:var(--navy)}@media(max-width:680px){main{padding:0 14px;margin-top:24px}.brand{padding:16px}.brand img{width:140px}.brand-label{font-size:9px;letter-spacing:1px}.brand-label strong{font-size:13px}h1{font-size:27px}.pad{padding:20px}.grid{grid-template-columns:1fr 1fr;gap:18px}.actions{align-items:stretch;flex-direction:column}.topline{align-items:flex-start;flex-direction:column}.success{padding:30px 20px}.total{gap:16px}.total strong{font-size:21px}}
        </style></head><body><header><div class="brand"><img src="${esc(CONFIG.logoUrl)}" alt="Molisana Imports"><div class="brand-label"><strong>Purchase order confirmation</strong></div></div></header><main>${body}</main><footer>Molisana Imports &bull; Vendor purchase order response</footer></body></html>`;
    }
    function statusPage(title, message, poNumber, success) {
        return page(`<section class="card success"><div class="icon">${success ? '&#10003;' : 'i'}</div><div class="eyebrow">${poNumber ? `Purchase order ${esc(poNumber)}` : 'Vendor portal'}</div><h1>${esc(title)}</h1><p>${esc(message)}</p><p class="hint">You may now close this window.</p></section>`);
    }
    function details(po, action, id, token, notes = '', error = '') {
        const approved = action === 'approved';
        const cell = (label, val, cls = '') => `<div><span class="label">${esc(label)}</span><div class="val ${cls}">${esc(val || '—')}</div></div>`;
        const rows = [];
        for (let line = 0, count = po.getLineCount({sublistId:'item'}); line < count; line++) {
            const get = fieldId => po.getSublistValue({sublistId:'item', fieldId, line});
            let item;
            try { item = po.getSublistText({sublistId:'item',fieldId:'item',line}); } catch (_) { item = get('item'); }
            rows.push(`<tr><td><strong>${esc(item)}</strong><div class="description">${esc(get('description'))}</div></td><td class="num">${esc(get('quantity'))}</td><td class="num">${esc(money(get('rate')))}</td><td class="num">${esc(money(get('amount')))}</td></tr>`);
        }
        return page(`<div class="eyebrow">Purchase order response</div><h1>Review your purchase order</h1><p class="sub">Review the details below and submit your ${approved ? 'acceptance' : 'rejection'}. You can add a note for our team.</p>
        <section class="card pad"><div class="topline"><h2 style="margin:0">Purchase Order ${esc(value(po,'tranid'))}</h2><span class="badge">Awaiting your response</span></div><div class="grid">${cell('Vendor',text(po,'entity'))}${cell('Order date',value(po,'trandate'))}${cell('Delivery due',date(value(po,'duedate')))}${cell('Currency',text(po,'currency'))}${cell('Payment terms',text(po,'terms'))}${cell('Ship to',value(po,'shipaddress'),'address')}</div></section>
        <section class="card"><div class="pad"><h2 style="margin:0">Order items</h2></div><div class="scroll"><table><thead><tr><th>Item / Description</th><th class="num">Quantity</th><th class="num">Rate</th><th class="num">Amount</th></tr></thead><tbody>${rows.join('') || '<tr><td colspan="4">No item lines on this purchase order.</td></tr>'}</tbody></table></div><div class="total"><span>PO total &middot; ${esc(text(po,'currency'))}</span><strong>${esc(money(value(po,'total')))}</strong></div><div class="hint" style="padding:0 26px 16px">PO total includes applicable charges and taxes recorded on the order.</div></section>
        <section class="card pad"><div class="topline"><h2 style="margin:0">Your response</h2><span class="badge ${approved ? 'accept' : 'reject'}">${approved ? 'Accept purchase order' : 'Reject purchase order'}</span></div>
        <form method="post">${error ? `<p class="notice" role="alert">${esc(error)}</p>` : ''}<input type="hidden" name="custpage_po_id" value="${esc(id)}"><input type="hidden" name="custpage_action" value="${esc(action)}"><input type="hidden" name="custpage_token" value="${esc(token)}"><input type="hidden" name="custpage_confirm" value="yes"><label for="notes">Notes <span>(optional)</span></label><textarea id="notes" name="custpage_notes" maxlength="${CONFIG.maxNotes}" placeholder="Add delivery details, questions, or a reason for your response…">${esc(notes)}</textarea><p class="hint">Up to ${CONFIG.maxNotes} characters. Your note will be shared with the Molisana team.</p><div class="actions"><p>Submitting will record this purchase order as ${approved ? 'accepted' : 'rejected'}. Contact Molisana if you need to change your response afterward.</p><button class="${approved ? '' : 'danger'}" type="submit">${approved ? 'Submit acceptance' : 'Submit rejection'}</button></div></form></section>`);
    }
    function onRequest(context) {
        const {request, response} = context;
        response.setHeader({name:'Content-Type', value:'text/html; charset=UTF-8'});
        response.setHeader({name:'Cache-Control', value:'no-store, max-age=0'});
        response.setHeader({name:'Referrer-Policy', value:'no-referrer'});
        response.setHeader({name:'X-Content-Type-Options', value:'nosniff'});
        response.setHeader({name:'Content-Security-Policy', value:"default-src 'none'; img-src https:; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'"});
        try {
            if (!['GET','POST'].includes(request.method)) fail('This request method is not supported.');
            const p = request.parameters;
            const post = request.method === 'POST';
            // id is supported for existing emails; custparam_po_id avoids NetSuite's reserved id parameter.
            const id = String(post ? p.custpage_po_id || '' : p.custparam_po_id || p.id || '');
            const action = String(post ? p.custpage_action || '' : p.action || '').trim().toLowerCase();
            const token = String(post ? p.custpage_token || '' : p.custparam_token || '');
            if (!/^[1-9]\d*$/.test(id) || !['approved','rejected'].includes(action)) fail('This link is incomplete or invalid. Please use the approval or rejection link in your email.');
            if (!/^\d{13}\.[a-f0-9]{64}$/.test(token)) fail('This link is missing its security token. Please contact Molisana for a new email link.');
            const po = record.load({type:record.Type.PURCHASE_ORDER, id, isDynamic:false});
           
            const current = String(value(po, CONFIG.decisionField) || '');
            if ([CONFIG.acceptedValue, CONFIG.rejectedValue].includes(current)) {
                // Block both reopening (GET) and resubmission (POST), including an old open form.
                response.write(statusPage('This link is closed', 'A decision has already been recorded for this purchase order. Further access and submissions are blocked. Please contact Molisana if a correction is needed.', '', false));
                return;
            }
            if (!post) { response.write(details(po, action, id, token)); return; }
            if (p.custpage_confirm !== 'yes') fail('Please open your email link and submit the confirmation form.');
            const notes = String(p.custpage_notes || '').trim();
            if (notes.length > CONFIG.maxNotes) {
                response.write(details(po, action, id, token, notes, `Please shorten your note to ${CONFIG.maxNotes} characters.`)); return;
            }
            po.setValue({fieldId: CONFIG.decisionField, value:action === 'approved' ? CONFIG.acceptedValue : CONFIG.rejectedValue});
            // Blank optional notes preserve existing notes. Memo and native approvalstatus are never set.
            if (notes) po.setValue({fieldId:CONFIG.notesField, value:notes});
            po.save({enableSourcing:false, ignoreMandatoryFields:false});
            response.write(statusPage('Thank you for your response', `Purchase order ${value(po,'tranid')} has been ${action === 'approved' ? 'accepted' : 'rejected'}.${notes ? ' Your notes have also been saved.' : ''}`, value(po,'tranid'), true));
        } catch (e) {
            // Do not log the request, token, vendor notes or full exception details.
            log.error({title:'Vendor PO response failed',details:String(e.name || 'ERROR')});
            const message = e.publicMessage || (e.name === 'RCRD_HAS_BEEN_CHANGED'
                ? 'This purchase order changed while you were responding. Reopen the email link to review its latest status.'
                : 'We could not complete your request. Please contact Molisana or try your email link again.');
            response.write(statusPage('Unable to complete your request', message, '', false));
        }
    }
    return {onRequest};
});
