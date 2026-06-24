/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 * @NModuleScope SameAccount
 */
define(['N/record', 'N/log'], function (record, log) {

    var CUSTOMER_DELIVERY_DAY_FIELD = 'custentity_mi_default_delivery_day';
    var TRANDATE_FIELD = 'trandate';
    var TARGET_FIELD = 'custbody_expected_shipping_date';

    function afterSubmit(context) {
        try {
            if (context.type !== context.UserEventType.EDIT && context.type !== context.UserEventType.CREATE) {
                return;
            }

            var rec = context.newRecord;
          
            var tranDate = rec.getValue({ fieldId: TRANDATE_FIELD });
            var custID = rec.getValue({ fieldId: 'entity' });
            var Cust_loaded = record.load({ type: 'customer', id: custID, isDynamic: false });
            var deliveryDayValue = Cust_loaded.getValue({ fieldId: CUSTOMER_DELIVERY_DAY_FIELD });

            var baseDate = toDate(tranDate);
            if (!baseDate) {
                log.debug({
                    title: 'Expected Shipping Date',
                    details: 'No usable trandate found. trandate=' + tranDate
                });
                return;
            }

            // --- Decide which rule applies -----------------
            var days = deliveryDayValueToDays(deliveryDayValue);
            var expectedDate;

            if (days.length > 0) {
                // Path 1: next upcoming day from trandate
                expectedDate = nextMatchingDay(baseDate, days);
            } else {
                // Path 2: trandate + 2 business days (weekends not counted)
                expectedDate = addLeadTime(baseDate);
            }

            // --- Persist ---------------------------------------------------
            var values = {};
            values[TARGET_FIELD] = expectedDate;

            record.submitFields({
                type: rec.type,
                id: rec.id,
                values: values,
                options: { enableSourcing: false, ignoreMandatoryFields: true }
            });

            log.audit({
                title: 'Expected Shipping Date set',
                details: 'SO ' + rec.id +
                         ' | deliveryDayValue="' + (deliveryDayValue || '') + '"' +
                         ' | days=[' + days.join(',') + ']' +
                         ' | trandate=' + baseDate +
                         ' -> ' + expectedDate
            });

        } catch (e) {
            log.error({
                title: 'Error populating ' + TARGET_FIELD,
                details: (e.name || '') + ' : ' + (e.message || e)
            });
        }
    }

    // ---------------------------------------------------------------------
    // Helpers
    // ---------------------------------------------------------------------

    /** Normalizes a date-field value to a Date object. */
    function toDate(value) {
        if (!value) { return null; }
        if (value instanceof Date) { return new Date(value.getTime()); }
        var d = new Date(value);
        return isNaN(d.getTime()) ? null : d;
    }

    /**
     * Returns the next date strictly AFTER fromDate whose weekday is in
     * dayNumbers. The earliest such day within the next 7 days wins.
     */
    function nextMatchingDay(fromDate, dayNumbers) {
        var result = new Date(fromDate.getTime());
        for (var i = 0; i < 7; i++) {
            result.setDate(result.getDate() + 1);
            if (dayNumbers.indexOf(result.getDay()) !== -1) {
                return result;
            }
        }
        return result; // safety fallback (shouldn't be reached)
    }

    function deliveryDayValueToDays(value) {
    var map = {
        '1': [1],    // Monday
        '2': [2],    // Tuesday
        '3': [3],    // Wednesday
        '4': [4],    // Thursday
        '5': [5],    // Friday
        '6': [1, 4]  // Monday/Thursday
    };

    return map[String(value || '')] || [];
}

    /**
     * Adds 2 business days (48 working hours) to startDate by stepping forward
     * one day at a time and counting only weekdays (Saturday and Sunday are
     * skipped and do not count toward the total).
     *
     * Examples:
     *   Mon -> Wed        Thu -> Mon (skip weekend)
     *   Tue -> Thu        Fri -> Tue (skip weekend)
     *   Wed -> Fri        Sat -> Tue   Sun -> Tue
     */
    function addLeadTime(startDate) {
        var result = new Date(startDate.getTime());
        var added = 0;
        while (added < 2) {
            result.setDate(result.getDate() + 1);
            var dow = result.getDay();
            if (dow !== 0 && dow !== 6) { // not Sunday and not Saturday
                added++;
            }
        }
        return result;
    }

    return { afterSubmit: afterSubmit };
});