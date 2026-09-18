/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 *
 * Simple purchase order approval page.
 * Open with: ...scriptlet.nl?script=###&deploy=###&poid=<internal id>
 */
define(['N/record', 'N/search'], function (record, search) {

    var NOTES_FIELD = 'custbody_po_approval_notes';   // or use 'memo'

    function onRequest(context) {
        var poId = context.request.parameters.poid;

        if (context.request.method === 'GET') {
            context.response.write(formPage(poId));
        } else {
            var action = context.request.parameters.action;
            var notes = context.request.parameters.notes || '';

            var values = {};
            values.approvalstatus = (action === 'approve') ? '2' : '3';
            values[NOTES_FIELD] = notes;

            record.submitFields({
                type: record.Type.PURCHASE_ORDER,
                id: poId,
                values: values,
                options: { ignoreMandatoryFields: true }
            });

            context.response.write(donePage(action === 'approve'));
        }
    }

    function formPage(poId) {
        var tranid = '';
        if (poId) {
            var f = search.lookupFields({
                type: search.Type.PURCHASE_ORDER,
                id: poId,
                columns: ['tranid']
            });
            tranid = f.tranid || '';
        }

        return wrap('Purchase order ' + esc(tranid),
            '<form method="POST">' +
            '<input type="hidden" name="poid" value="' + esc(poId) + '">' +
            '<input type="hidden" name="action" id="action">' +
            '<button type="button" class="btn ok" onclick="openNotes(\'approve\')">Approve</button> ' +
            '<button type="button" class="btn no" onclick="openNotes(\'reject\')">Reject</button>' +
            '<div id="box">' +
            '<div class="modal">' +
            '<h2>Add notes</h2>' +
            '<textarea name="notes" id="notes" rows="5" ' +
            'placeholder="Write anything you want the buyer to see"></textarea>' +
            '<div class="row">' +
            '<button type="button" class="btn" onclick="document.getElementById(\'box\').style.display=\'none\'">Cancel</button> ' +
            '<button type="submit" class="btn ok">Submit</button>' +
            '</div></div></div>' +
            '</form>' +
            '<script>function openNotes(a){' +
            'document.getElementById("action").value=a;' +
            'document.getElementById("box").style.display="flex";' +
            'document.getElementById("notes").focus();}<\/script>');
    }

    function donePage(approved) {
        return wrap(approved ? 'Approved' : 'Rejected',
            '<p>Thank you. Your response has been saved.</p>');
    }

    function wrap(title, body) {
        return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
            '<meta name="viewport" content="width=device-width,initial-scale=1">' +
            '<title>' + esc(title) + '</title><style>' +
            'body{font-family:Arial,sans-serif;background:#f2f4f7;margin:0;font-size:14px}' +
            '.wrap{max-width:520px;margin:40px auto;background:#fff;border:1px solid #d5d9de;padding:26px}' +
            'h1{margin:0 0 20px;font-size:20px}' +
            '.btn{padding:9px 20px;font-size:14px;border:1px solid #b6bdc5;background:#eef0f3;cursor:pointer}' +
            '.btn.ok{background:#2f7d32;border-color:#2f7d32;color:#fff}' +
            '.btn.no{background:#b3261e;border-color:#b3261e;color:#fff}' +
            '#box{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);' +
            'align-items:center;justify-content:center}' +
            '.modal{background:#fff;padding:22px;width:90%;max-width:400px}' +
            '.modal h2{margin:0 0 12px;font-size:17px}' +
            'textarea{width:100%;box-sizing:border-box;padding:8px;border:1px solid #b6bdc5;font:inherit}' +
            '.row{margin-top:14px;text-align:right}' +
            '</style></head><body><div class="wrap"><h1>' + esc(title) + '</h1>' +
            body + '</div></body></html>';
    }

    function esc(v) {
        return String(v == null ? '' : v)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;')
            .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    return { onRequest: onRequest };
});