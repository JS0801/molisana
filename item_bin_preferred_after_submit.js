/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 */
define(['N/record', 'N/log'], (record, log) => {
  const LOCATION_ID = 315;
  const BIN_SUBLIST_ID = 'binnumber';

  const afterSubmit = (context) => {
    if (context.type === context.UserEventType.DELETE) {
      return;
    }

    const itemRec = record.load({
      type: context.newRecord.type,
      id: context.newRecord.id,
      isDynamic: false
    });

    const lineCount = itemRec.getLineCount({ sublistId: BIN_SUBLIST_ID });
    const preferredLines = [];
    let changed = false;

    logBinNumberSublist(itemRec, context.newRecord.id, 'Bin number sublist before preferred update');

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
      const onHand = itemRec.getSublistValue({
        sublistId: BIN_SUBLIST_ID,
        fieldId: 'onhand',
        line: i
      });
      const isPreferred = itemRec.getSublistValue({
        sublistId: BIN_SUBLIST_ID,
        fieldId: 'preferredbin',
        line: i
      });

      if (
        String(lineLocation) === String(LOCATION_ID) &&
        isPositive(onHand) &&
        !isChecked(isPreferred)
      ) {
        itemRec.setSublistValue({
          sublistId: BIN_SUBLIST_ID,
          fieldId: 'preferredbin',
          line: i,
          value: true
        });

        preferredLines.push({
          line: i,
          locationId: lineLocation,
          binId: lineBin,
          onHand,
          oldPreferred: isPreferred,
          newPreferred: true
        });

        changed = true;
      }
    }

    if (!changed) {
      log.audit('No preferred bin changes needed', `Item ${context.newRecord.id}`);
      return;
    }

    const savedId = itemRec.save({
      enableSourcing: true,
      ignoreMandatoryFields: true
    });

    log.audit('Preferred bin lines checked', JSON.stringify({
      itemId: savedId,
      locationId: LOCATION_ID,
      preferredLines
    }));
    logBinNumberSublist(itemRec, savedId, 'Bin number sublist after preferred update');
  };

  const isPositive = (value) => {
    return Number(value || 0) > 0;
  };

  const isChecked = (value) => {
    return value === true || value === 'T';
  };

  const logBinNumberSublist = (itemRec, itemId, title) => {
    const lineCount = itemRec.getLineCount({ sublistId: BIN_SUBLIST_ID });
    const lines = [];

    for (let i = 0; i < lineCount; i++) {
      lines.push({
        line: i,
        locationId: itemRec.getSublistValue({
          sublistId: BIN_SUBLIST_ID,
          fieldId: 'location',
          line: i
        }),
        binId: itemRec.getSublistValue({
          sublistId: BIN_SUBLIST_ID,
          fieldId: 'binnumber',
          line: i
        }),
        onHand: itemRec.getSublistValue({
          sublistId: BIN_SUBLIST_ID,
          fieldId: 'onhand',
          line: i
        }),
        preferred: itemRec.getSublistValue({
          sublistId: BIN_SUBLIST_ID,
          fieldId: 'preferredbin',
          line: i
        })
      });
    }

    log.audit(title, JSON.stringify({
      itemId,
      lineCount,
      lines
    }));
  };

  return { afterSubmit };
});
