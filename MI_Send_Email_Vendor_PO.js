/**
 * @NApiVersion 2.1
 * @NScriptType WorkflowActionScript
 */
define(['N/record', 'N/email', 'N/config', 'N/format', 'N/log', 'N/render', 'N/url', 'N/crypto/random'],
function (record, email, config, format, log, render, url, random) {

    // Vendor response Suitelet deployment supplied for this account.
    // This file is self-contained; no separate email link helper is required.
    var SETTINGS = {
        suiteletScriptId: 'customscript_mi_po_approval_from_vendor',
        suiteletDeploymentId: 'customdeploy_mi_po_approval_from_vendor',
        decisionField: 'custbody_mi_vendor_decision',
        tokenField: 'custbody_mi_vendor_link_token',
        linkValidDays: 14,
        logoUrl: 'https://4975346-sb1.app.netsuite.com/core/media/media.nl?id=4770&c=4975346_SB1&h=sGvHCgrcrHMJzoZjmKF-9Og7Y_nEEncIbLigVYoLtPhS1CXd'
    };

    function escapeHtml(value) {
        return String(value == null ? '' : value).replace(/[&<>"']/g, function (ch) {
            return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[ch];
        });
    }

    function buildLinks(poRec) {
        if (SETTINGS.suiteletScriptId.indexOf('REPLACE') !== -1 ||
            SETTINGS.suiteletDeploymentId.indexOf('REPLACE') !== -1) {
            throw new Error('Configure the vendor response Suitelet script and deployment IDs in SETTINGS.');
        }
        var storedToken = String(poRec.getValue({fieldId: SETTINGS.tokenField}) || '');
        // Reuse an unexpired token so reminder emails do not invalidate earlier links.
        var token = storedToken;
        if (!/^\d{13}\.[a-f0-9]{64}$/.test(token) || Number(token.split('.')[0]) <= Date.now()) {
            token = String(Date.now() + SETTINGS.linkValidDays * 86400000) + '.' +
                Array.from(random.generateBytes({size: 32}), function (b) {
                    return b.toString(16).padStart(2, '0');
                }).join('');
        }
        function resolve(action) {
            return url.resolveScript({
                scriptId: SETTINGS.suiteletScriptId,
                deploymentId: SETTINGS.suiteletDeploymentId,
                returnExternalUrl: true,
                // Suitelet accepts this alias. NetSuite reserves the plain "id" parameter.
                params: {custparam_po_id: String(poRec.id), action: action, custparam_token: token}
            });
        }
        var links = {approved: resolve('approved'), rejected: resolve('rejected')};
        if (token !== storedToken) {
            record.submitFields({
                type: record.Type.PURCHASE_ORDER,
                id: poRec.id,
                values: {[SETTINGS.tokenField]: token},
                options: {enableSourcing: false, ignoreMandatoryFields: false}
            });
        }
        return links;
    }

    function onAction(scriptContext) {
        try {
            // Configure this action AFTER RECORD SUBMIT, excluding XEDIT.
            // Token persistence uses submitFields (XEDIT); do not send again for that update.
            if (String(scriptContext.type || '').toLowerCase() === 'xedit') return;
            var poId = scriptContext.newRecord.id;
            if (!poId) throw new Error('Run this email action after the purchase order has been saved.');
            var poRec = record.load({type: record.Type.PURCHASE_ORDER, id: poId});
            var decision = String(poRec.getValue({fieldId: SETTINGS.decisionField}) || '');
            if (decision === '1' || decision === '2') {
                log.debug('PO Email', 'Approval email skipped: PO ' + poId + ' already has a vendor decision.');
                return;
            }

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

            // Render the existing PO attachment before persisting any new link token.
            var pdfFile = render.transaction({
                entityId: Number(poId),
                printMode: render.PrintMode.PDF,
                inCustLocale: true
            });
            pdfFile.name = 'Purchase Order ' + tranId + '.pdf';
            var links = buildLinks(poRec);
            var currency = poRec.getText({fieldId: 'currency'}) || '';
            var amount = Number(total).toLocaleString('en-US', {
                minimumFractionDigits: 2, maximumFractionDigits: 2
            });
            function row(label, content) {
                return '<tr><td style="padding:10px 14px;border-bottom:1px solid #e8eaf0;color:#687386;width:40%;vertical-align:top;">' +
                    escapeHtml(label) + '</td><td style="padding:10px 14px;border-bottom:1px solid #e8eaf0;color:#172749;">' +
                    escapeHtml(content).replace(/\r?\n/g, '<br>') + '</td></tr>';
            }
            function button(label, href, background) {
                return '<td style="padding:0 12px 12px 0;"><table role="presentation" border="0" cellspacing="0" cellpadding="0"><tr><td bgcolor="' + background +
                    '" style="border-radius:6px;mso-padding-alt:14px 22px;text-align:center;"><a href="' + escapeHtml(href) +
                    '" style="display:inline-block;padding:14px 22px;color:#ffffff;text-decoration:none;font-size:14px;font-weight:bold;border:1px solid ' + background +
                    ';border-radius:6px;">' + escapeHtml(label) + '</a></td></tr></table></td>';
            }

            // Email-safe table layout with inline styles; no JavaScript in the email.
            var body = '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
                '<body style="margin:0;padding:0;background:#f4f3ef;font-family:Arial,Helvetica,sans-serif;color:#26344b;">' +
                '<table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#f4f3ef"><tr><td align="center" style="padding:28px 12px;">' +
                '<table role="presentation" width="640" border="0" cellspacing="0" cellpadding="0" style="width:100%;max-width:640px;background:#ffffff;border:1px solid #e4e7ec;">' +
                '<tr><td style="padding:24px 28px;border-top:5px solid #172749;border-bottom:1px solid #e4e7ec;">' +
                '<img src="' + escapeHtml(SETTINGS.logoUrl) + '" width="170" alt="Molisana Imports" style="display:block;width:170px;height:auto;border:0;"></td></tr>' +
                '<tr><td style="padding:28px;"><p style="margin:0 0 8px;color:#927630;font-size:11px;font-weight:bold;letter-spacing:2px;">PURCHASE ORDER CONFIRMATION</p>' +
                '<h1 style="margin:0 0 20px;color:#172749;font-size:26px;line-height:1.25;">Purchase Order ' + escapeHtml(tranId) + '</h1>' +
                '<p style="font-size:14px;line-height:1.7;">Dear ' + escapeHtml(entityName) + ',</p>' +
                '<p style="font-size:14px;line-height:1.7;">Please review the purchase order below and confirm whether you can accept it. The full purchase order is attached as a PDF.</p>' +
                '<table width="100%" border="0" cellspacing="0" cellpadding="0" style="font-size:13px;background:#fafbfc;border:1px solid #e8eaf0;">' +
                row('PO Number', tranId) + row('Date', trandateStr) + row('Vendor', entityName) +
                row('Amount', currency + ' ' + amount) + row('Requested Delivery Date', duedateStr) + row('Ship To', shipAddress) + '</table>' +
                '<h2 style="margin:26px 0 10px;font-size:18px;color:#172749;">Confirm your response</h2>' +
                '<p style="font-size:14px;line-height:1.7;">Choose an option below to open our vendor portal, review the PO, and add optional notes. Your decision is recorded only after you click Submit on that page.</p>' +
                '<table role="presentation" border="0" cellspacing="0" cellpadding="0"><tr>' +
                button('Approve PO', links.approved, '#172749') + button('Reject PO', links.rejected, '#a63c31') + '</tr></table>' +
                '<p style="font-size:12px;line-height:1.6;color:#697586;">Once submitted, the response link will close. Please contact our purchasing team if you need to change your decision.</p>' +
                '<p style="font-size:12px;line-height:1.6;color:#697586;">If the buttons do not work, use these links: <a href="' + escapeHtml(links.approved) +
                '" style="color:#172749;">Approve PO</a> &nbsp;|&nbsp; <a href="' + escapeHtml(links.rejected) + '" style="color:#a63c31;">Reject PO</a>.</p>' +
                '<p style="font-size:14px;line-height:1.7;margin-top:24px;">If you have questions or foresee any issues meeting the delivery date, please reply to this email or contact our purchasing team.</p>' +
                '<p style="font-size:14px;line-height:1.7;">Best regards,<br><strong>' + escapeHtml(companyName) + '</strong><br>' + escapeHtml(mainPhone) + '</p></td></tr>' +
                '<tr><td style="padding:16px 28px;background:#f8f7f3;color:#697586;font-size:11px;">' + escapeHtml(companyName) +
                ' &bull; Please keep your response links private.</td></tr></table></td></tr></table></body></html>';

            // -----------------------------
            // 6. Send Email
            // -----------------------------
            email.send({
                author: 12425,                 // Sender: Purchasing Molisana
                recipients: [12138],//recipients,     // Array of email addresses
                subject: subject,
                body: body,
                attachments: [pdfFile],
                relatedRecords: {
                    transactionId: poId
                }
            });

            log.debug('PO Email', 'Email sent for PO ' + poId + ' to: ' + recipients.join(', '));

        } catch (e) {
            // Avoid writing private email links or vendor notes to the execution log.
            log.error('Error sending PO email', {name: e.name || 'ERROR', message: e.message || 'Unable to send PO email'});
        }
    }

    return {
        onAction: onAction
    };
});
