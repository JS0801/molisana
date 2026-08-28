/*eslint-disable*/

var FINAL_APPROVER_ID = '-5'//'12138';

var VALIDATED_FIELD = 'custbody_vendbill_validator_check';
var NOTES_FIELD     = 'custbody_mi_approval_notes';

function process(email) {
    try {
        // Subject: Approve_VENDBILL48768 / Reject_VENDBILL48768
        var match = String(email.getSubject() || '').trim().match(/^(Approve|Reject)_(.+)$/i);
        if (!match) return;

        var action = match[1].toLowerCase();
        var billNumber = match[2].trim();
        var emailBody = String(email.getTextBody() || '');

        // Find Vendor Bill + transaction currency amount
        var results = nlapiSearchRecord('vendorbill', null,
            [
                ['type', 'anyof', 'VendBill'], 'AND',
                ['transactionnumber', 'is', billNumber], 'AND',
                ['mainline', 'is', 'T']
            ],
            [
                new nlobjSearchColumn('internalid'),
                new nlobjSearchColumn('fxamount')
            ]
        );

        if (!results || !results.length) {
            nlapiLogExecution('ERROR', 'Bill Not Found', billNumber);
            return;
        }

        var billId = results[0].getId();
        var amount = Math.abs(parseFloat(results[0].getValue('fxamount') || 0) || 0);

        // Load bill so workflow sees a normal record save
        var bill = nlapiLoadRecord('vendorbill', billId);
        var validated = bill.getFieldValue(VALIDATED_FIELD) === 'T';
        var oldNotes = bill.getFieldValue(NOTES_FIELD) || '';

        // Take only whatever is after COMMENTS:
        var commentPos = emailBody.toUpperCase().indexOf('COMMENTS:');
        var comment = commentPos >= 0 ? emailBody.substring(commentPos + 9).trim() : '';

        // Append comment only when a comment exists
        if (comment) {
            bill.setFieldValue(
                NOTES_FIELD,
                oldNotes +
                (oldNotes ? '\n\n--------------------\n\n' : '') +
                nlapiDateToString(new Date()) + '\n' +
                comment
            );
        }

        nlapiLogExecution('AUDIT', 'Bill',
            billNumber + ' | Amount: ' + amount +
            ' | Validated: ' + validated +
            ' | Action: ' + action
        );

        // Reject - sender does not matter
        if (action === 'reject') {
            bill.setFieldValue('approvalstatus', '3');
            nlapiSubmitRecord(bill, true, false);
            return;
        }

        // Get sender email
        var from = email.getFrom();
        var senderEmail = from ? String(from.getEmail() || '').trim().toLowerCase() : '';

        if (!senderEmail) {
            nlapiLogExecution('ERROR', 'Sender Email Missing', billNumber);
            return;
        }

        // Find employee from sender email
        var employees = nlapiSearchRecord('employee', null,
            [
                ['email', 'is', senderEmail], 'AND',
                ['isinactive', 'is', 'F']
            ],
            [new nlobjSearchColumn('internalid')]
        );

        if (!employees || !employees.length) {
            nlapiLogExecution('ERROR', 'Employee Not Found', senderEmail);
            return;
        }

        var senderId = String(employees[0].getId());

        nlapiLogExecution('AUDIT', 'Email Sender',
            'Email: ' + senderEmail + ' | Employee ID: ' + senderId
        );

        // Employee -5 = final approver
        // Make sure Validated = T and approve regardless of amount
        if (senderId === FINAL_APPROVER_ID) {
            bill.setFieldValue(VALIDATED_FIELD, 'T');
            bill.setFieldValue('approvalstatus', '2');
            nlapiSubmitRecord(bill, true, false);
            return;
        }

        // Any other employee = first approval
        bill.setFieldValue(VALIDATED_FIELD, 'T');

        // <= $1,00 = approve
        if (amount <= 100) {
            bill.setFieldValue('approvalstatus', '2');
        }

        // > $1,000 = Validated only; remains Pending Approval
        nlapiSubmitRecord(bill, true, false);

    } catch (e) {
        nlapiLogExecution('ERROR', 'Email Capture Error', e.toString());
    }
}