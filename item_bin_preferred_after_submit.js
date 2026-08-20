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
    const removedLines = [];
    let changed = false;

    logBinNumberSublist(itemRec, context.newRecord.id, 'Bin number sublist before cleanup');

    for (let i = lineCount - 1; i >= 0; i--) {
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
        isZero(onHand) &&
        !isChecked(isPreferred)
      ) {
        removedLines.push({
          line: i,
          locationId: lineLocation,
          binId: lineBin,
          onHand,
          preferred: isPreferred
        });

        itemRec.removeLine({
          sublistId: BIN_SUBLIST_ID,
          line: i,
          ignoreRecalc: true
        });
        changed = true;
      }
    }

    if (!changed) {
      log.audit('No bin number lines removed', `Item ${context.newRecord.id}`);
      return;
    }

    const savedId = itemRec.save({
      enableSourcing: true,
      ignoreMandatoryFields: true
    });

    log.audit('Removed bin number lines', JSON.stringify({
      itemId: savedId,
      locationId: LOCATION_ID,
      removedLines
    }));
    logBinNumberSublist(itemRec, savedId, 'Bin number sublist after cleanup');
  };

  const isZero = (value) => {
    return Number(value || 0) === 0;
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
