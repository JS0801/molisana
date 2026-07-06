/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 */
define(['N/runtime', 'N/email', 'N/log'], function (runtime, email, log) {

  function esc(s) {
    s = (s == null ? '' : String(s));
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function urlEncode(s) {
    return encodeURIComponent(String(s || ''));
  }

  function afterSubmit(context) {
    try {
      if (context.type !== context.UserEventType.EDIT) return;

      var rec = context.newRecord;
      var poId = rec.id;
      if (!poId) return;

      var tranid = rec.getValue({ fieldId: 'tranid' }) || ('PO ' + poId);

      // You can send to vendor email or internal user. Example: send to vendor email field (entity) is not available on newRecord easily.
      // For now using current user as you had hardcoded recipients:
      var senderId = 12138;
      var recipientId = 12138;

      var replyToEmail = 'emails.4975346.2753.98b7678d92@4975346.email.netsuite.com';

      // Add token so inbound email script can find PO
      var subject = 'PO Acknowledgement Required | ' + tranid + ' | [POID:' + poId + ']';

      var mailtoSubject = subject;
      var mailtoBody =
        'Hi,%0D%0A%0D%0A' +
        'Please find attached Proforma Invoice (mandatory) as proof of acknowledgement for ' + tranid + '.%0D%0A%0D%0A' +
        'Thank you.%0D%0A';

      var mailtoHref =
        'mailto:' + replyToEmail +
        '?subject=' + urlEncode(mailtoSubject) +
        '&body=' + mailtoBody;

      var htmlBody = ''
        + '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.45;">'
        + '  <p>Hi,</p>'
        + '  <p><b>Action required:</b> Please upload/send the <b>Proforma Invoice</b> and acknowledge receipt for Purchase Order <b>' + esc(tranid) + '</b>.</p>'
        + '  <p style="margin:10px 0 16px 0;">'
        + '    <span style="font-size:18px;">📎</span> <b>Proforma invoice attachment is mandatory</b> (proof of acknowledgement).'
        + '  </p>'

        // Button
        + '  <p style="margin:18px 0;">'
        + '    <a href="' + esc(mailtoHref) + '" '
        + '       style="display:inline-block;padding:12px 18px;border-radius:10px;'
        + '              background:#2563eb;color:#fff;text-decoration:none;font-weight:bold;">'
        + '      Reply with Invoice'
        + '    </a>'
        + '  </p>'

        // Backup
        + '  <p style="color:#666;font-size:12px;">'
        + '    If the button does not work, please reply to this email and attach the Proforma Invoice. '
        + '    Make sure the subject includes: <b>[POID:' + esc(poId) + ']</b>'
        + '  </p>'

        + '  <p>Thanks</p>'
        + '</div>';

      email.send({
        author: senderId,
        recipients: [recipientId],
        subject: subject,
        body: htmlBody,
        replyTo: replyToEmail
      });

    } catch (e) {
      log.error('afterSubmit error', e);
    }
  }

  return { afterSubmit: afterSubmit };
});