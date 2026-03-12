/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 * @NModuleScope SameAccount
 */
define(['N/search', 'N/record', 'N/log'], (search, record, log) => {

  // Map item "Type" text from search to record.Type constants
  function mapItemTypeToRecordType(typeText, itemlot) {
    log.debug('itemlot', itemlot)
    
    
    if (itemlot === 'T') {
        typeText = 'Lot-numbered ' + typeText;
      }
    log.debug('typeText', typeText)
    var itemTypeLookup = {
        'Assembly': record.Type.ASSEMBLY_ITEM,
        'Inventory Item': record.Type.INVENTORY_ITEM,
        'Lot-numbered Inventory Item': record.Type.LOT_NUMBERED_INVENTORY_ITEM,
        'Lot-numbered InvtPart': record.Type.LOT_NUMBERED_INVENTORY_ITEM,
        'Lot-numbered Assembly Item': record.Type.LOT_NUMBERED_ASSEMBLY_ITEM,
        'Lot-numbered Assembly': record.Type.LOT_NUMBERED_ASSEMBLY_ITEM,
        'Serialized Inventory Item': record.Type.SERIALIZED_INVENTORY_ITEM,
        'Serialized Assembly Item': record.Type.SERIALIZED_INVENTORY_ITEM,
        'Non-inventory Item': record.Type.NON_INVENTORY_ITEM,
        'Discount': record.Type.DISCOUNT_ITEM,
        'Other Charge': record.Type.OTHER_CHARGE_ITEM,
        'Service': record.Type.SERVICE_ITEM,
        'Description': record.Type.DESCRIPTION_ITEM,
        'Kit/Package': record.Type.KIT_ITEM,
        'Item Group': record.Type.ITEM_GROUP,
        'InvtPart': record.Type.INVENTORY_ITEM
      };
     log.debug('typeTextFinal', itemTypeLookup[typeText])

    // fallback: treat as inventory item (adjust if your account needs different default)
    return itemTypeLookup[typeText];
  }

  function getInputData() {
    // Build the exact search you provided (grouped at item level)
var itemSearchObj = search.create({
   type: "item",
   filters:
   [
      ["transaction.type","anyof","CustInvc"],
      "AND", 
      ["transaction.mainline","is","F"], 
      "AND", 
      ["formulatext: case when {transaction.locationnohierarchy} = {inventorylocation} then 1 else 0 end","is","1"], 
      "AND", 
      ["isinactive","is","F"], 
      "AND", 
      [["transaction.trandate","within","lastyear"],"OR",["transaction.trandate","within","thisyeartodate"]], 
      "AND", 
      [
        ["max(formulanumeric: CASE WHEN SUM(NVL(CASE        WHEN TO_CHAR({transaction.trandate}, 'YYYY') = TO_CHAR(ADD_MONTHS(SYSDATE, -12), 'YYYY')     AND {transaction.locationnohierarchy} = {inventorylocation}       THEN NVL({transaction.quantity}, 0)        ELSE 0  END, 0)) <> MAX(NVL({custitem_qty_sold_last_year}, 0)) THEN 1 ELSE 0 END)","equalto","1"],
        "OR",
        ["max(formulanumeric: CASE WHEN SUM(NVL(CASE       WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 30 AND TRUNC(SYSDATE)     AND {transaction.locationnohierarchy} = {inventorylocation}      THEN NVL({transaction.quantity},0)       ELSE 0 END, 0)) <> MAX(NVL({custitem_qty_sold_in_last_30_days}, 0)) THEN 1 ELSE 0 END)","equalto","1"],
        "OR",
        ["max(formulanumeric: CASE WHEN SUM(NVL(CASE       WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 60 AND TRUNC(SYSDATE)  -31   AND {transaction.locationnohierarchy} = {inventorylocation}    THEN NVL({transaction.quantity},0)       ELSE 0 END, 0)) <> MAX(NVL({custitem_qty_sold_in_last_60_days}, 0)) THEN 1 ELSE 0 END)","equalto","1"],
        "OR",
        ["max(formulanumeric: CASE WHEN SUM(NVL(CASE       WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 90 AND TRUNC(SYSDATE)    -61   AND {transaction.locationnohierarchy} = {inventorylocation}    THEN NVL({transaction.quantity},0)       ELSE 0 END, 0)) <> MAX(NVL({custitem_qty_sold_in_last_90_days}, 0)) THEN 1 ELSE 0 END)","equalto","1"],
        "OR",
        ["max(formulanumeric: CASE WHEN SUM(NVL(CASE       WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 120 AND TRUNC(SYSDATE)    -91    AND {transaction.locationnohierarchy} = {inventorylocation}   THEN NVL({transaction.quantity},0)       ELSE 0 END, 0)) <> MAX(NVL({custitem_qty_sold_in_last_120_days}, 0)) THEN 1 ELSE 0 END)","equalto","1"],
        "OR",
        ["max(formulanumeric: CASE WHEN SUM(NVL(CASE       WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 120 AND TRUNC(SYSDATE)   AND {transaction.locationnohierarchy} = {inventorylocation}    THEN NVL({transaction.quantity},0)       ELSE 0 END, 0)) <> MAX(NVL({custitem_qty_sold_in_4_months}, 0)) THEN 1 ELSE 0 END)","equalto","1"]
      ]
   ],
   columns:
   [
      search.createColumn({
         name: "internalid",
         summary: "GROUP",
         label: "Item ID"
      }),
      search.createColumn({
         name: "itemid",
         summary: "GROUP",
         label: "Item Name"
      }),
      search.createColumn({
         name: "islotitem",
         summary: "GROUP",
         label: "Is Lot Numbered Item"
      }),
      search.createColumn({
         name: "type",
         summary: "GROUP",
         label: "Type"
      }),
      search.createColumn({
         name: "formulanumeric",
         summary: "SUM",
         formula: "CASE     WHEN TO_CHAR({transaction.trandate}, 'YYYY') = TO_CHAR(ADD_MONTHS(SYSDATE, -12), 'YYYY')     AND {transaction.locationnohierarchy} = {inventorylocation}    THEN NVL({transaction.quantity}, 0)     ELSE 0  END",
         label: "Previous Year QTY Sold"
      }),
      search.createColumn({
         name: "formulanumeric30",
         summary: "SUM",
         formula: "CASE     WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 30 AND TRUNC(SYSDATE)    AND {transaction.locationnohierarchy} = {inventorylocation}     THEN NVL({transaction.quantity},0)     ELSE 0  END",
         label: "Last 30 Days"
      }),
      search.createColumn({
         name: "formulanumeric60",
         summary: "SUM",
         formula: "CASE     WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 60 AND TRUNC(SYSDATE)  -31   AND {transaction.locationnohierarchy} = {inventorylocation}    THEN NVL({transaction.quantity},0)     ELSE 0  END",
         label: "Last 60 Days"
      }),
      search.createColumn({
         name: "formulanumeric90",
         summary: "SUM",
         formula: "CASE     WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 90 AND TRUNC(SYSDATE)  -61    AND {transaction.locationnohierarchy} = {inventorylocation}   THEN NVL({transaction.quantity},0)     ELSE 0  END",
         label: "Last 90 Days"
      }),
      search.createColumn({
         name: "formulanumeric120",
         summary: "SUM",
         formula: "CASE     WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 120 AND TRUNC(SYSDATE)  -91    AND {transaction.locationnohierarchy} = {inventorylocation}   THEN NVL({transaction.quantity},0)     ELSE 0  END",
         label: "Last 120 Days"
      }),
      search.createColumn({
         name: "formulanumeric4",
         summary: "SUM",
         formula: "CASE     WHEN {transaction.trandate} BETWEEN TRUNC(SYSDATE) - 120 AND TRUNC(SYSDATE)   AND {transaction.locationnohierarchy} = {inventorylocation}    THEN NVL({transaction.quantity},0)     ELSE 0  END",
         label: "Last 4 Months"
      })
   ]
});

    return itemSearchObj;
  }

  function map(context) {
    try {
      var result = JSON.parse(context.value);

      log.debug('result', result)

      var itemId = result.values['GROUP(internalid)']; // internal id
      var itemTypeText = result.values['GROUP(type)'];  // e.g., "Inventory Item", "Assembly/Bill of Materials"
      var currentFieldStr = result.values['MAX(custitem_qty_sold_last_year)']; // current field value
      var itemlot = result.values['GROUP(islotitem)'];

      var sumQtyLastYr = result.values['SUM(formulanumeric)'];
      var sumQtyLast30 = result.values['SUM(formulanumeric30)'];
      var sumQtyLast60 = result.values['SUM(formulanumeric60)'];
      var sumQtyLast90 = result.values['SUM(formulanumeric90)'];
      var sumQtyLast120 = result.values['SUM(formulanumeric120)'];
      var sumQtyLast4 = result.values['SUM(formulanumeric4)'];

      

      var recType = mapItemTypeToRecordType(itemTypeText.value, itemlot);
      // Fast update with submitFields
      record.submitFields({
        type: recType,
        id: itemId.value,
        values: { 
          custitem_qty_sold_last_year: sumQtyLastYr,
          custitem_qty_sold_in_last_30_days: sumQtyLast30,
          custitem_qty_sold_in_last_60_days: sumQtyLast60,
          custitem_qty_sold_in_last_90_days: sumQtyLast90,
          custitem_qty_sold_in_last_120_days: sumQtyLast120,
          custitem_qty_sold_in_4_months: sumQtyLast4
        },
        options: { enablesourcing: false, ignoreMandatoryFields: true }
      });

    } catch (e) {
      log.error('MAP Error', e);
    }
  }

  function reduce(context) {}

  function summarize(summary) {}

  return { getInputData, map, reduce, summarize };
});
