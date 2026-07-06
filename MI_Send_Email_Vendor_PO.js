/**
 * @NApiVersion 2.x
 * @NScriptType WorkflowActionScript
 */
define(['N/record', 'N/email', 'N/config', 'N/format', 'N/log', 'N/render'],
function (record, email, config, format, log, render) {

    function onAction(scriptContext) {
        try {
            var poRec = scriptContext.newRecord;
            var poId  = poRec.id;

            // -----------------------------
            // 1. Get Vendor + Email List
            // -----------------------------
            var vendorId = poRec.getValue({ fieldId: 'entity' });
            if (!vendorId) {
                log.debug('PO Email', 'No vendor on PO ' + poId);
                return;
            }

            var vendorRec = record.load({
                type: record.Type.VENDOR,
                id: vendorId
            });

            var emailField = vendorRec.getValue({
                fieldId: 'custentity_po_email_addresses'
            }) || '';

            emailField = String(emailField);

            var recipients = [];
            if (emailField) {
                emailField.split(';').forEach(function (addr) {
                    if (!addr) return;
                    var clean = addr.trim();
                    if (clean) {
                        recipients.push(clean);
                    }
                });
            }

            if (!recipients.length) {
                log.debug('PO Email', 'No email addresses in custentity_po_email_addresses for vendor ' + vendorId);
                return;
            }

            // -----------------------------
            // 2. Get PO Fields
            // -----------------------------
            var tranId      = poRec.getValue({ fieldId: 'tranid' }) || '';
            var trandate    = poRec.getValue({ fieldId: 'trandate' });
            var duedate     = poRec.getValue({ fieldId: 'duedate' });
            var entityName  = poRec.getText({ fieldId: 'entity' }) || '';
            var total       = poRec.getValue({ fieldId: 'total' }) || 0;
            var shipAddress = poRec.getValue({ fieldId: 'shipaddress' }) || '';

            // Format dates
            var trandateStr = '';
            var duedateStr  = '';

            if (trandate) {
                trandateStr = format.format({
                    value: trandate,
                    type: format.Type.DATE
                });
            }

            if (duedate) {
                duedateStr = format.format({
                    value: duedate,
                    type: format.Type.DATE
                });
            }

            // -----------------------------
            // 3. Company Info
            // -----------------------------
            var companyConfig = config.load({
                type: config.Type.COMPANY_INFORMATION
            });

            var companyName = companyConfig.getValue({ fieldId: 'companyname' }) || '';
            var mainPhone   = companyConfig.getValue({ fieldId: 'phone' }) || '';

            // -----------------------------
            // 4. Subject
            // -----------------------------
            var subject = 'New Purchase Order ' + tranId + ' from ' + companyName;

            // -----------------------------
            // 5. HTML Body
            // -----------------------------
            var body = ''
                + '<p>'
                + '  Dear ' + entityName + ', We are pleased to share a new Purchase Order for your review and processing.'
                + '  <br />'
                + '  &nbsp;'
                + '</p>'
                + '<p><strong>Purchase Order Details</strong></p>'
                + '<figure class="table">'
                + '  <table border="0" cellpadding="2" cellspacing="0">'
                + '    <tbody>'
                + '      <tr>'
                + '        <td><strong>PO Number:</strong></td>'
                + '        <td>' + tranId + '</td>'
                + '      </tr>'
                + '      <tr>'
                + '        <td><strong>Date:</strong></td>'
                + '        <td>' + trandateStr + '</td>'
                + '      </tr>'
                + '      <tr>'
                + '        <td><strong>Vendor:</strong></td>'
                + '        <td>' + entityName + '</td>'
                + '      </tr>'
                + '      <tr>'
                + '        <td><strong>Amount:</strong></td>'
                + '        <td>' + total + '</td>'
                + '      </tr>'
                + '      <tr>'
                + '        <td><strong>Delivery Date (Requested):</strong></td>'
                + '        <td>' + duedateStr + '</td>'
                + '      </tr>'
                + '      <tr>'
                + '        <td><strong>Ship To:</strong></td>'
                + '        <td>' + shipAddress + '</td>'
                + '      </tr>'
                + '    </tbody>'
                + '  </table>'
                + '</figure>'
                + '<p>'
                + '  <br />'
                + '  You can view the full details of this Purchase Order in the attached document or via your usual process. '
                + 'If you have any questions or foresee any issues meeting the requested delivery date, please reply to this email or contact our purchasing team.'
                + '  <br /><br />'
                + '  Best regards,<br />'
                +    companyName + '<br />'
                +    mainPhone
                + '</p>';


          var pdfFile = render.transaction({
                entityId: poId,
                printMode: render.PrintMode.PDF,
                inCustLocale: true
            });

            pdfFile.name = 'Purchase Order ' + tranId + '.pdf';

            // -----------------------------
            // 6. Send Email
            // -----------------------------
            email.send({
                author: 12425,                 // Sender: Purchasing Molisana
                recipients: recipients,     // Array of email addresses
                subject: subject,
                cc: ['Fabio@gruppodb.ca','qc@molisana.com'],
                body: body,
                attachments: [pdfFile],
                relatedRecords: {
                    transactionId: poId
                }
            });

            log.debug('PO Email', 'Email sent for PO ' + poId + ' to: ' + recipients.join(', '));

        } catch (e) {
            log.error('Error sending PO email', e);
        }
    }

    return {
        onAction: onAction
    };
});
