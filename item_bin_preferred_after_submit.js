/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 */
define(['N/record', 'N/log'], (record, log) => {
  const BIN_FIELD_ID = 'custitem_bin';
  const LOCATION_ID = 315;
  const BIN_SUBLIST_ID = 'binnumber';

  const afterSubmit = (context) => {
    if (context.type === context.UserEventType.DELETE) {
      return;
    }

    const binId = context.newRecord.getValue({ fieldId: BIN_FIELD_ID });

    if (!binId) {
      return;
    }

    const itemRec = record.load({
      type: context.newRecord.type,
      id: context.newRecord.id,
      isDynamic: false
    });

    const lineCount = itemRec.getLineCount({ sublistId: BIN_SUBLIST_ID });
    let foundLine = -1;
    let changed = false;

    for (let i = 0; i < lineCount; i++) {
      const lineLocation = itemRec.getSublistValue({
        sublistId: BIN_SUBLIST_ID,
        fieldId: 'location',
        line: i
      });
      const lineBin = itemRec.getSublistValue({
        sublistId: BIN_SUBLIST_ID,
        fieldId: 'binnumber',
        line: i
      });
      const isPreferred = itemRec.getSublistValue({
        sublistId: BIN_SUBLIST_ID,
        fieldId: 'preferredbin',
        line: i
      });

      if (String(lineLocation) !== String(LOCATION_ID)) {
        continue;
      }

      if (String(lineBin) === String(binId)) {
        foundLine = i;

        if (!isPreferred) {
          itemRec.setSublistValue({
            sublistId: BIN_SUBLIST_ID,
            fieldId: 'preferredbin',
            line: i,
            value: true
          });
          changed = true;
        }
      } else if (isPreferred) {
        itemRec.setSublistValue({
          sublistId: BIN_SUBLIST_ID,
          fieldId: 'preferredbin',
          line: i,
          value: false
        });
        changed = true;
      }
    }

    if (foundLine === -1) {
      itemRec.setSublistValue({
        sublistId: BIN_SUBLIST_ID,
        fieldId: 'location',
        line: lineCount,
        value: LOCATION_ID
      });
      itemRec.setSublistValue({
        sublistId: BIN_SUBLIST_ID,
        fieldId: 'binnumber',
        line: lineCount,
        value: binId
      });
      itemRec.setSublistValue({
        sublistId: BIN_SUBLIST_ID,
        fieldId: 'preferredbin',
        line: lineCount,
        value: true
      });
      changed = true;
    }

    if (!changed) {
      return;
    }

    const savedId = itemRec.save({
      enableSourcing: true,
      ignoreMandatoryFields: true
    });

    log.audit('Preferred bin updated', `Item ${savedId}, bin ${binId}, location ${LOCATION_ID}`);
  };

  return { afterSubmit };
});
