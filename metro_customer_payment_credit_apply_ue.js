/**
 * Metro Customer Payment Credit Apply - User Event
 *
 * @NApiVersion 2.x
 * @NScriptType UserEventScript
 */
define(['N/search', 'N/log'], function(search, log) {
    var CREDIT_JSON_FIELD_ID = 'custbody_metro_credit_apply_json'; // TODO: replace with the Payment custom field mapped from CSV column "Notes".

    function beforeSubmit(context) {
        if (context.type !== context.UserEventType.CREATE &&
                context.type !== context.UserEventType.EDIT) {
            return;
        }

        var payment = context.newRecord;
        var paymentId = payment.id || '(new customer payment)';
        var jsonText = payment.getValue({ fieldId: CREDIT_JSON_FIELD_ID });

        log.debug({
            title: 'Payment credit apply beforeSubmit',
            details: 'type=' + context.type +
                ' paymentId=' + paymentId +
                ' jsonPresent=' + (!!jsonText) +
                ' jsonLength=' + (jsonText ? String(jsonText).length : 0)
        });

        if (!jsonText) {
            log.debug({
                title: 'No credit apply JSON found',
                details: 'paymentId=' + paymentId + ' fieldId=' + CREDIT_JSON_FIELD_ID
            });
            return;
        }

        var payload;
        try {
            payload = JSON.parse(jsonText);
        } catch (e) {
            log.error({
                title: 'Invalid credit apply JSON',
                details: 'paymentId=' + paymentId +
                    ' error=' + getErrorDetails(e) +
                    ' value=' + jsonText
            });
            return;
        }

        log.debug({
            title: 'Credit apply JSON parsed',
            details: 'paymentId=' + paymentId +
                ' sourceFile=' + ((payload && payload.fileName) || '') +
                ' paymentNumber=' + ((payload && payload.paymentNumber) || '') +
                ' creditCount=' + ((payload && payload.credits && payload.credits.length) || 0)
        });

        if (!payload || !payload.credits || !payload.credits.length) {
            log.audit({
                title: 'No credits in JSON payload',
                details: 'paymentId=' + paymentId
            });
            return;
        }

        try {
            applyCreditsToPayment(payment, paymentId, payload);
        } catch (e) {
            log.error({
                title: 'Credit apply failed',
                details: 'paymentId=' + paymentId + ' error=' + getErrorDetails(e)
            });
        }
    }

    function applyCreditsToPayment(payment, paymentId, payload) {
        var customerId = getBodyValue(payment, ['customer', 'entity']);
        var creditLineCount = payment.getLineCount({ sublistId: 'credit' }) || 0;
        var operations = [];
        var errors = [];

        log.debug({
            title: 'Credit apply setup',
            details: 'paymentId=' + paymentId +
                ' customerId=' + customerId +
                ' creditSublistLineCount=' + creditLineCount +
                ' payloadCreditCount=' + payload.credits.length
        });

        for (var i = 0; i < payload.credits.length; i++) {
            var credit = payload.credits[i] || {};
            var amount = parsePositiveAmount(credit.amount);
            var creditId = resolveCreditInternalId(credit, customerId);

            log.debug({
                title: 'Credit apply row check',
                details: 'paymentId=' + paymentId +
                    ' rowNumber=' + (credit.rowNumber || '') +
                    ' creditTranId=' + (credit.creditTranId || '') +
                    ' creditInternalIdFromJson=' + (credit.creditInternalId || '') +
                    ' resolvedCreditId=' + creditId +
                    ' rawAmount=' + credit.amount +
                    ' parsedAmount=' + amount
            });

            if (!creditId) {
                errors.push('Credit row ' + credit.rowNumber + ' has no credit internal id or searchable credit tranid.');
                continue;
            }

            if (!(amount > 0)) {
                errors.push('Credit ' + creditId + ' has invalid amount: ' + credit.amount);
                continue;
            }

            var line = findSublistLineByValue(payment, 'credit', 'doc', String(creditId));
            log.debug({
                title: 'Credit sublist line search',
                details: 'paymentId=' + paymentId +
                    ' creditId=' + creditId +
                    ' foundLine=' + line +
                    ' creditSublistLineCount=' + creditLineCount
            });

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

            log.debug({
                title: 'Applying credit line',
                details: 'paymentId=' + paymentId +
                    ' creditId=' + operation.creditId +
                    ' line=' + operation.line +
                    ' amount=' + operation.amount
            });

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

        log.debug({
            title: 'Credit apply JSON cleared',
            details: 'paymentId=' + paymentId + ' fieldId=' + CREDIT_JSON_FIELD_ID
        });

        log.audit({
            title: 'Credits applied to customer payment',
            details: 'paymentId=' + paymentId + ' creditCount=' + operations.length + ' sourceFile=' + (payload.fileName || '')
        });
    }

    function resolveCreditInternalId(credit, customerId) {
        if (credit.creditInternalId) {
            log.debug({
                title: 'Credit internal id provided',
                details: 'creditInternalId=' + credit.creditInternalId +
                    ' creditTranId=' + (credit.creditTranId || '')
            });
            return String(credit.creditInternalId);
        }

        if (!credit.creditTranId) {
            log.debug({
                title: 'Credit lookup skipped',
                details: 'No creditTranId found on payload row ' + (credit.rowNumber || '')
            });
            return '';
        }

        var filters = [
            ['tranid', 'is', String(credit.creditTranId)]
        ];

        if (customerId) {
            filters.push('AND');
            filters.push(['entity', 'anyof', customerId]);
        }

        log.debug({
            title: 'Searching credit memo',
            details: 'creditTranId=' + credit.creditTranId + ' customerId=' + (customerId || '')
        });

        var resultSet = search.create({
            type: search.Type.CREDIT_MEMO,
            filters: filters,
            columns: ['internalid']
        }).run().getRange({
            start: 0,
            end: 1
        });

        var creditId = resultSet && resultSet.length ? String(resultSet[0].getValue({ name: 'internalid' })) : '';

        log.debug({
            title: 'Credit memo search result',
            details: 'creditTranId=' + credit.creditTranId + ' resolvedCreditId=' + creditId
        });

        return creditId;
    }

    function findSublistLineByValue(rec, sublistId, fieldId, expectedValue) {
        var line = rec.findSublistLineWithValue({
            sublistId: sublistId,
            fieldId: fieldId,
            value: expectedValue
        });

        if (line < 0 && /^\d+$/.test(String(expectedValue))) {
            line = rec.findSublistLineWithValue({
                sublistId: sublistId,
                fieldId: fieldId,
                value: parseInt(expectedValue, 10)
            });
        }

        return line;
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
        beforeSubmit: beforeSubmit
    };
});
