/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 * @NModuleScope SameAccount
 *
 * Sales Order - afterSubmit
 * Populates custbody_expected_shipping_date.
 *
 * Logic:
 *   custbody_delivery_note is a FREE-TEXT field that may contain day codes,
 *   e.g. "OBU056 | FRI/MON".
 *
 *   1. If one or more day codes are found in custbody_delivery_note:
 *        expected date = the next upcoming date AFTER trandate that matches
 *        one of those days (the earliest matching day wins).
 *
 *   2. If no day code is found (field empty or no recognizable day):
 *        expected date = trandate + 2 business days (48 working hours),
 *        counting weekdays only -- Saturday and Sunday are skipped and do
 *        not count toward the 48 hours.
 *
 * Notes:
 *   - "Upcoming" is treated as strictly AFTER trandate. If trandate itself
 *     falls on a matching day, the NEXT occurrence is used. To make it
 *     on-or-after instead, change the loop start in nextMatchingDay().
 *   - In afterSubmit, context.newRecord can read back empty for some saves,
 *     so we fall back to record.load. submitFields persists the value and
 *     does NOT re-trigger this user event (no recursion).
 */
define(['N/record', 'N/log'], function (record, log) {

    var DELIVERY_NOTE_FIELD = 'custbody_delivery_note';
    var TRANDATE_FIELD = 'trandate';
    var TARGET_FIELD = 'custbody_expected_shipping_date';

    // Day name / abbreviation -> day-of-week number (0 = Sun ... 6 = Sat)
    var DAY_MAP = {
        SUN: 0, SUNDAY: 0,
        MON: 1, MONDAY: 1,
        TUE: 2, TUES: 2, TUESDAY: 2,
        WED: 3, WEDS: 3, WEDNESDAY: 3,
        THU: 4, THUR: 4, THURS: 4, THURSDAY: 4,
        FRI: 5, FRIDAY: 5,
        SAT: 6, SATURDAY: 6
    };

    function afterSubmit(context) {
        try {
            if (context.type !== context.UserEventType.EDIT && context.type !== context.UserEventType.CREATE) {
                return;
            }

            var rec = context.newRecord;
          
            var tranDate = rec.getValue({ fieldId: TRANDATE_FIELD });
            var custID = rec.getValue({ fieldId: 'entity' });
            var Cust_loaded = record.load({ type: 'customer', id: custID, isDynamic: false });
            var deliveryNote = Cust_loaded.getValue({ fieldId: 'custentity2' });

            var baseDate = toDate(tranDate);
            if (!baseDate) {
                log.debug({
                    title: 'Expected Shipping Date',
                    details: 'No usable trandate found. trandate=' + tranDate
                });
                return;
            }

            // --- Decide which rule applies -----------------
            var days = extractDays(deliveryNote);
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
                         ' | deliveryNote="' + (deliveryNote || '') + '"' +
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
     * Extracts day-of-week numbers from the free-text delivery note.
     * Prefers the segment after the last "|" (e.g. "OBU056 | FRI/MON" -> "FRI/MON")
     * to avoid matching letters inside a route code. Falls back to the whole
     * string if nothing is found there.
     * Returns an array like [5, 1] for "FRI/MON" (order preserved, de-duped).
     */
    function extractDays(text) {
        if (!text) { return []; }
        var s = String(text);

        var segment = (s.indexOf('|') !== -1)
            ? s.substring(s.lastIndexOf('|') + 1)
            : s;

        var found = scanDays(segment);
        if (found.length === 0 && segment !== s) {
            found = scanDays(s); // fallback: scan the whole string
        }
        return found;
    }

    function scanDays(str) {
        var found = [];
        var tokens = String(str).toUpperCase().match(/[A-Z]+/g) || [];
        for (var i = 0; i < tokens.length; i++) {
            var tok = tokens[i];
            if (DAY_MAP.hasOwnProperty(tok)) {
                var num = DAY_MAP[tok];
                if (found.indexOf(num) === -1) { found.push(num); }
            }
        }
        return found;
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