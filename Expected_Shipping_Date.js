/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 * @NModuleScope SameAccount
 *
 * Sales Order - Expected Shipping Date & Sales Channel
 *
 * CREATE : Calculate header Expected Shipping Date from Customer delivery day
 *          (or 2 business days lead time), set Sales Channel from Customer,
 *          and push the date to every item line.
 * EDIT   : If the header Expected Shipping Date changed, push it to every item line.
 *          If unchanged, do nothing.
 *
 * Runs in beforeSubmit, so everything is written in the same save.
 * No record.load, no submitFields, no second save.
 */
define(['N/search', 'N/log'], (search, log) => {

    const CUSTOMER_DELIVERY_DAY_FIELD = 'custentity_mi_default_delivery_day';
    const CUSTOMER_SALES_CHANNEL_FIELD = 'custentity_default_shipping_channel';
    const TRANDATE_FIELD = 'trandate';
    const TARGET_FIELD = 'custbody_expected_shipping_date';
    const SO_SALES_CHANNEL_FIELD = 'saleschannel';
    const LINE_FIELD = 'expectedshipdate';
    const LEAD_TIME_BUSINESS_DAYS = 2;

    // Delivery day list value -> JS weekday numbers (0 = Sun ... 6 = Sat)
    const DELIVERY_DAY_MAP = {
        '1': [1],
        '2': [2],
        '3': [3],
        '4': [4],
        '5': [5],
        '6': [1, 4]
    };

    const beforeSubmit = (context) => {
        const UE = context.UserEventType;
        if (context.type !== UE.CREATE && context.type !== UE.EDIT) return;

        try {
            const rec = context.newRecord;
            let expectedDate;

            if (context.type === UE.CREATE) {
                const baseDate = toDate(rec.getValue({ fieldId: TRANDATE_FIELD }));
                const custId = rec.getValue({ fieldId: 'entity' });

                if (!baseDate || !custId) {
                    log.debug('Expected Shipping Date', `Skipped. trandate=${baseDate} | customer=${custId}`);
                    return;
                }

                // One lightweight lookup instead of loading the full customer record
                const cust = search.lookupFields({
                    type: search.Type.CUSTOMER,
                    id: custId,
                    columns: [CUSTOMER_DELIVERY_DAY_FIELD, CUSTOMER_SALES_CHANNEL_FIELD]
                });

                const deliveryDay = selectValue(cust[CUSTOMER_DELIVERY_DAY_FIELD]);
                const salesChannel = selectValue(cust[CUSTOMER_SALES_CHANNEL_FIELD]);
                const days = DELIVERY_DAY_MAP[deliveryDay] || [];

                expectedDate = days.length
                    ? nextMatchingDay(baseDate, days)
                    : addBusinessDays(baseDate, LEAD_TIME_BUSINESS_DAYS);

                // Header fields - both set in the same save
                rec.setValue({ fieldId: TARGET_FIELD, value: expectedDate });
                if (salesChannel) {
                    rec.setValue({ fieldId: SO_SALES_CHANNEL_FIELD, value: salesChannel });
                }

                log.audit('Header set', `Customer ${custId} | deliveryDay="${deliveryDay}" | days=[${days}] | ` +
                    `trandate=${baseDate.toDateString()} -> ${expectedDate.toDateString()} | salesChannel=${salesChannel}`);

            } else {
                // EDIT - only act when the header date actually changed
                expectedDate = toDate(rec.getValue({ fieldId: TARGET_FIELD }));
                const oldDate = toDate(context.oldRecord.getValue({ fieldId: TARGET_FIELD }));

                if (!expectedDate || sameDay(oldDate, expectedDate)) return;
            }

            const updated = setLineDates(rec, expectedDate);

            log.audit('Line Expected Ship Date updated',
                `SO ${rec.id || '(new)'} | ${context.type} | Date: ${expectedDate.toDateString()} | Lines changed: ${updated}`);

        } catch (e) {
            log.error(`Error populating ${TARGET_FIELD}`, `${e.name || ''} : ${e.message || e}`);
        }
    };

    // Sets the line date only where it differs; returns number of lines changed
    const setLineDates = (rec, date) => {
        const lineCount = rec.getLineCount({ sublistId: 'item' });
        let changed = 0;

        for (let i = 0; i < lineCount; i++) {
            const current = toDate(rec.getSublistValue({ sublistId: 'item', fieldId: LINE_FIELD, line: i }));
            if (sameDay(current, date)) continue;

            rec.setSublistValue({ sublistId: 'item', fieldId: LINE_FIELD, line: i, value: date });
            changed++;
        }
        return changed;
    };

    // lookupFields returns select fields as [{ value, text }]
    const selectValue = (v) => {
        if (Array.isArray(v)) return v.length ? String(v[0].value) : '';
        return v ? String(v) : '';
    };

    const toDate = (value) => {
        if (!value) return null;
        const d = value instanceof Date ? new Date(value.getTime()) : new Date(value);
        return isNaN(d.getTime()) ? null : d;
    };

    const sameDay = (a, b) => !!a && !!b && a.toDateString() === b.toDateString();

    const nextMatchingDay = (fromDate, dayNumbers) => {
        const result = new Date(fromDate.getTime());
        for (let i = 0; i < 7; i++) {
            result.setDate(result.getDate() + 1);
            if (dayNumbers.includes(result.getDay())) return result;
        }
        return result;
    };

    const addBusinessDays = (startDate, count) => {
        const result = new Date(startDate.getTime());
        let added = 0;
        while (added < count) {
            result.setDate(result.getDate() + 1);
            const dow = result.getDay();
            if (dow !== 0 && dow !== 6) added++;
        }
        return result;
    };

    return { beforeSubmit };
});