/**
 * @NApiVersion 2.x
 * @NScriptType UserEventScript
 *
 * Vendor Bill afterSubmit automation.
 *
 * On Vendor Bill create:
 *   1. Read custbody_note_to_vendor from the bill
 *   2. Find a Vendor Credit whose tranid matches that value
 *   3. Load the Vendor Credit
 *   4. Find the current bill on the Vendor Credit apply sublist
 *   5. Apply the bill to the credit
 *
 * Scope intentionally stays narrow:
 *   - create only
 *   - approved bills only
 *   - unpaid bills only
 *   - no Credit Memo automation
 */
define(['N/record', 'N/search', 'N/log'], function(record, search, log) {

    var BILL_NOTE_FIELD_ID = 'custbody_note_to_vendor';
    var APPROVED_STATUS_ID = '2';

    function afterSubmit(context) {
        if (context.type !== context.UserEventType.CREATE) {
            return;
        }

        var billId = context.newRecord.id;
        if (!billId) {
            log.error({
                title: 'Missing bill id',
                details: 'Vendor Bill afterSubmit did not expose context.newRecord.id.'
            });
            return;
        }

        try {
            var billInfo = getBillInfo(billId);
            if (!billInfo) {
                log.error({
                    title: 'Vendor Bill not found',
                    details: 'billId=' + billId
                });
                return;
            }

            if (!billInfo.noteToVendor) {
                log.audit({
                    title: 'No note to vendor',
                    details: 'billId=' + billId + ' has no ' + BILL_NOTE_FIELD_ID + '; skipping Vendor Credit application.'
                });
                return;
            }

            if (String(billInfo.approvalStatus) !== APPROVED_STATUS_ID) {
                log.audit({
                    title: 'Bill is not approved',
                    details: 'billId=' + billId + ' approvalstatus=' + billInfo.approvalStatus + '; skipping Vendor Credit application.'
                });
                return;
            }

            if (!(billInfo.amountRemaining > 0)) {
                log.audit({
                    title: 'Bill is not unpaid',
                    details: 'billId=' + billId + ' amountremaining=' + billInfo.amountRemaining + '; skipping Vendor Credit application.'
                });
                return;
            }

            var vendorCreditId = findVendorCreditByTranId(billInfo.noteToVendor);
            if (!vendorCreditId) {
                log.audit({
                    title: 'Matching Vendor Credit not found',
                    details: 'billId=' + billId + ' expected vendorcredit.tranid=' + billInfo.noteToVendor
                });
                return;
            }

            var applied = applyBillToVendorCredit(vendorCreditId, billId);
            if (!applied) {
                return;
            }

            log.audit({
                title: 'Bill applied to Vendor Credit',
                details: 'billId=' + billId + ' vendorCreditId=' + vendorCreditId + ' matchedTranId=' + billInfo.noteToVendor
            });

        } catch (e) {
            log.error({
                title: 'Vendor Credit application failed',
                details: 'billId=' + billId + ' error=' + getErrorDetails(e)
            });
        }
    }

    function getBillInfo(billId) {
        var billSearch = search.create({
            type: search.Type.VENDOR_BILL,
            filters: [
                ['internalid', 'anyof', billId],
                'AND',
                ['mainline', 'is', 'T']
            ],
            columns: [
                BILL_NOTE_FIELD_ID,
                'approvalstatus',
                'amountremaining'
            ]
        });

        var result = null;
        billSearch.run().each(function(row) {
            result = {
                noteToVendor: trim(row.getValue({ name: BILL_NOTE_FIELD_ID })),
                approvalStatus: row.getValue({ name: 'approvalstatus' }),
                amountRemaining: parseAmount(row.getValue({ name: 'amountremaining' }))
            };
            return false;
        });

        return result;
    }

    function findVendorCreditByTranId(tranId) {
        var creditId = null;

        var creditSearch = search.create({
            type: search.Type.VENDOR_CREDIT,
            filters: [
                ['tranid', 'is', tranId],
                'AND',
                ['mainline', 'is', 'T']
            ],
            columns: [
                search.createColumn({ name: 'internalid', sort: search.Sort.DESC })
            ]
        });

        creditSearch.run().each(function(row) {
            creditId = row.getValue({ name: 'internalid' });
            return false;
        });

        return creditId;
    }

    function applyBillToVendorCredit(vendorCreditId, billId) {
        var creditRecord = record.load({
            type: record.Type.VENDOR_CREDIT,
            id: vendorCreditId,
            isDynamic: false
        });

        var lineCount = creditRecord.getLineCount({ sublistId: 'apply' });
        var billIdText = String(billId);

        for (var line = 0; line < lineCount; line++) {
            var applyDocId = creditRecord.getSublistValue({
                sublistId: 'apply',
                fieldId: 'doc',
                line: line
            });

            if (String(applyDocId) !== billIdText) {
                continue;
            }

            creditRecord.setSublistValue({
                sublistId: 'apply',
                fieldId: 'apply',
                line: line,
                value: true
            });

            var dueAmount = creditRecord.getSublistValue({
                sublistId: 'apply',
                fieldId: 'due',
                line: line
            });

            if (dueAmount !== null && dueAmount !== undefined && dueAmount !== '') {
                creditRecord.setSublistValue({
                    sublistId: 'apply',
                    fieldId: 'amount',
                    line: line,
                    value: dueAmount
                });
            }

            creditRecord.save({
                enableSourcing: true,
                ignoreMandatoryFields: true
            });
            return true;
        }

        log.audit({
            title: 'Bill not present on Vendor Credit apply sublist',
            details: 'billId=' + billId + ' vendorCreditId=' + vendorCreditId
        });
        return false;
    }

    function parseAmount(value) {
        var amount = parseFloat(String(value || '').replace(/,/g, ''));
        return isNaN(amount) ? 0 : amount;
    }

    function trim(value) {
        return String(value || '').replace(/^\s+|\s+$/g, '');
    }

    function getErrorDetails(error) {
        if (!error) {
            return '';
        }

        if (error.name || error.message) {
            return (error.name || 'Error') + ': ' + (error.message || '');
        }

        return String(error);
    }

    return {
        afterSubmit: afterSubmit
    };
});
