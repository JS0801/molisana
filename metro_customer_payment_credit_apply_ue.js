/**
 * Metro Customer Payment Credit Apply - User Event
 *
 * @NApiVersion 2.x
 * @NScriptType UserEventScript
 */
define(['N/record', 'N/search', 'N/log'], function(record, search, log) {
    var CREDIT_JSON_FIELD_ID = 'custbody_metro_credit_apply_json'; // TODO: replace with the Payment custom field mapped from CSV column "Notes".

    function afterSubmit(context) {
        if (context.type !== context.UserEventType.CREATE &&
                context.type !== context.UserEventType.EDIT) {
            return;
        }

        var paymentId = context.newRecord.id;
        var jsonText = context.newRecord.getValue({ fieldId: CREDIT_JSON_FIELD_ID });
        if (!jsonText) {
            return;
        }

        var payload;
        try {
            payload = JSON.parse(jsonText);
        } catch (e) {
            log.error({
                title: 'Invalid credit apply JSON',
                details: 'paymentId=' + paymentId + ' error=' + getErrorDetails(e) + ' value=' + jsonText
            });
            return;
        }

        if (!payload || !payload.credits || !payload.credits.length) {
            log.audit({
                title: 'No credits in JSON payload',
                details: 'paymentId=' + paymentId
            });
            return;
        }

        try {
            applyCreditsToPayment(paymentId, payload);
        } catch (e) {
            log.error({
                title: 'Credit apply failed',
                details: 'paymentId=' + paymentId + ' error=' + getErrorDetails(e)
            });
        }
    }

    function applyCreditsToPayment(paymentId, payload) {
        var payment = record.load({
            type: record.Type.CUSTOMER_PAYMENT,
            id: paymentId,
            isDynamic: false
        });

        var customerId = getBodyValue(payment, ['customer', 'entity']);
        var operations = [];
        var errors = [];

        for (var i = 0; i < payload.credits.length; i++) {
            var credit = payload.credits[i] || {};
            var amount = parsePositiveAmount(credit.amount);
            var creditId = resolveCreditInternalId(credit, customerId);

            if (!creditId) {
                errors.push('Credit row ' + credit.rowNumber + ' has no credit internal id or searchable credit tranid.');
                continue;
            }

            if (!(amount > 0)) {
                errors.push('Credit ' + creditId + ' has invalid amount: ' + credit.amount);
                continue;
            }

            var line = findSublistLineByValue(payment, 'credit', 'doc', String(creditId));
            if (line < 0) {
                errors.push('Credit ' + creditId + ' was not found on payment credit sublist.');
                continue;
            }

            operations.push({
                line: line,
                creditId: creditId,
                amount: amount
            });
        }

        if (errors.length) {
            log.error({
                title: 'Credit apply validation failed',
                details: 'paymentId=' + paymentId + ' errors=' + errors.join(' | ')
            });
            return;
        }

        for (var opIndex = 0; opIndex < operations.length; opIndex++) {
            var operation = operations[opIndex];
            payment.setSublistValue({
                sublistId: 'credit',
                fieldId: 'apply',
                line: operation.line,
                value: true
            });
            payment.setSublistValue({
                sublistId: 'credit',
                fieldId: 'amount',
                line: operation.line,
                value: operation.amount
            });
        }

        payment.setValue({
            fieldId: CREDIT_JSON_FIELD_ID,
            value: ''
        });

        var savedId = payment.save({
            enableSourcing: false,
            ignoreMandatoryFields: true
        });

        log.audit({
            title: 'Credits applied to customer payment',
            details: 'paymentId=' + savedId + ' creditCount=' + operations.length + ' sourceFile=' + (payload.fileName || '')
        });
    }

    function resolveCreditInternalId(credit, customerId) {
        if (credit.creditInternalId) {
            return String(credit.creditInternalId);
        }

        if (!credit.creditTranId) {
            return '';
        }

        var filters = [
            ['tranid', 'is', String(credit.creditTranId)]
        ];

        if (customerId) {
            filters.push('AND');
            filters.push(['entity', 'anyof', customerId]);
        }

        var resultSet = search.create({
            type: search.Type.CREDIT_MEMO,
            filters: filters,
            columns: ['internalid']
        }).run().getRange({
            start: 0,
            end: 1
        });

        return resultSet && resultSet.length ? String(resultSet[0].getValue({ name: 'internalid' })) : '';
    }

    function findSublistLineByValue(rec, sublistId, fieldId, expectedValue) {
        var lineCount = rec.getLineCount({ sublistId: sublistId }) || 0;
        for (var i = 0; i < lineCount; i++) {
            var value = rec.getSublistValue({
                sublistId: sublistId,
                fieldId: fieldId,
                line: i
            });

            if (String(value) === String(expectedValue)) {
                return i;
            }
        }

        return -1;
    }

    function getBodyValue(rec, fieldIds) {
        for (var i = 0; i < fieldIds.length; i++) {
            try {
                var value = rec.getValue({ fieldId: fieldIds[i] });
                if (value) {
                    return value;
                }
            } catch (e) {
                // Try the next possible field id.
            }
        }

        return '';
    }

    function parsePositiveAmount(value) {
        var text = String(value || '').replace(/[,$\s]/g, '');
        if (!/^\d+(\.\d+)?$/.test(text)) {
            return 0;
        }

        return parseFloat(text);
    }

    function getErrorDetails(error) {
        if (!error) {
            return '';
        }

        return [
            error.name || '',
            error.message || '',
            error.stack || ''
        ].join(' | ');
    }

    return {
        afterSubmit: afterSubmit
    };
});
