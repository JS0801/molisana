/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define([
  'N/ui/serverWidget',
  'N/file',
  'N/log',
  'N/search',
  'N/record',
  'N/runtime',
  'N/crypto'
], function (ui, file, log, search, record, runtime, crypto) {

  var PORTAL_URL = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2110&deploy=1&compid=4975346&ns-at=AAEJ7tMQamzukv1WMqTK6i2c27bRetbrd2MDLjhDgPPFOawMxCo';

  var SOURCE_FOLDER_1 = 402335;
  var SOURCE_FOLDER_2 = 402334;
  var SOURCE_FOLDER_3 = 413248;

  var DOWNLOAD_FOLDER_UI = 378271;
  var DOWNLOAD_FOLDER_CRON = 279208;

  var TOKEN_TTL_MS = 30 * 60 * 1000;

  function onRequest(context) {
    try {
      if (context.request.method === 'GET') {
        handleGet(context);
      } else {
        handlePost(context);
      }
    } catch (e) {
      log.error('onRequest error', e);
      context.response.write('<html><body><h3>Unexpected error</h3><pre>' + escapeHtml(e.name + ': ' + e.message) + '</pre></body></html>');
    }
  }

  function handleGet(context) {
    var req = context.request;
    var q = req.parameters || {};

    var typeParam = String(q.type || '3');
    var showTopFilters = (typeParam !== '4');
    var formName = 'MI Reorder Tool (' + (typeParam === '4' ? 'Basic' : 'Admin') + ')';

    var mode = String(q.mode || '').toLowerCase();
    var cronTs = q.cronts || '';
    var cronSig = q.cronsig || '';
    var isCron = (mode === 'cron' && verifyCron(cronTs, cronSig));

    var form = ui.createForm({ title: formName });

    var empid = q.empid || '';
    var ts = q.ts || '';
    var sig = q.sig || '';
    var selectedEmp = q.custpage_id || '';

    if (empid && ts && sig && verifyUser(empid, ts, sig)) {
      selectedEmp = empid;
    }

    if (!selectedEmp && !isCron) {
      context.response.write(
        '<html><head>' +
        '<script>setTimeout(function(){ window.location.href = ' + JSON.stringify(PORTAL_URL) + '; }, 1200);</script>' +
        '<style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>' +
        '</head><body><div class="message">Login Required</div></body></html>'
      );
      return;
    }

    addHiddenFields(form, selectedEmp, ts, sig);

    var htmlField = form.addField({
      id: 'custpage_excel_html',
      type: ui.FieldType.INLINEHTML,
      label: 'Excel Table'
    });

    var selectedField = form.addField({
      id: 'custpage_selected',
      type: ui.FieldType.LONGTEXT,
      label: 'Selected Rows JSON'
    });
    selectedField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

    form.addSubmitButton({ label: 'Submit' });

    var latestFiles = getLatestFilesFromFolders();
    if (!latestFiles.file1) {
      htmlField.defaultValue = '<p style="color:red;">No source file found.</p>';
      context.response.writePage(form);
      return;
    }

    var sourceRows = buildMergedRows(latestFiles);
    if (!sourceRows || sourceRows.length <= 1) {
      htmlField.defaultValue = '<p style="color:red;">No valid data found in source files.</p>';
      context.response.writePage(form);
      return;
    }

    var headers = parseCsvLine(sourceRows[0]);
    var normalizedHeaders = [];
    var i;
    for (i = 0; i < headers.length; i++) {
      normalizedHeaders.push(normHeader(headers[i]));
    }

    var adminCsvIndex = normalizedHeaders.indexOf('admin portal');
    if (adminCsvIndex < 0) {
      for (i = 0; i < normalizedHeaders.length; i++) {
        if (
          normalizedHeaders[i] === 'admin' ||
          normalizedHeaders[i].indexOf('admin portal') !== -1
        ) {
          adminCsvIndex = i;
          break;
        }
      }
    }

    var truncateAfterAdmin = (typeParam === '4');
    var removeJustAdmin = (typeParam === '3');

    var processed = processRows({
      sourceRows: sourceRows,
      headers: headers,
      adminCsvIndex: adminCsvIndex,
      truncateAfterAdmin: truncateAfterAdmin,
      removeJustAdmin: removeJustAdmin,
      showTopFilters: showTopFilters
    });

    var finalCsv = buildCsvString(processed.newContent);
    var uiFileId = saveCsvFile('Download_MI_Reorder.csv', finalCsv, DOWNLOAD_FOLDER_UI, true);
    var cronFileId = saveCsvFile('RR Tool Details.csv', finalCsv, DOWNLOAD_FOLDER_CRON, true);
    var cronFileObj = file.load({ id: cronFileId });
    var dlUrl = cronFileObj.url;

    var fileField = form.addField({
      id: 'custpage_file_id',
      type: ui.FieldType.TEXT,
      label: 'File ID'
    });
    fileField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });
    fileField.defaultValue = String(uiFileId || '');

    if (isCron) {
      context.response.setHeader({ name: 'Content-Type', value: 'application/json' });
      context.response.write(JSON.stringify({
        ok: true,
        fileId: uiFileId,
        url: dlUrl,
        name: cronFileObj.name
      }));
      return;
    }

    htmlField.defaultValue = buildHtml({
      tableHeaders: processed.tableHeaders,
      rowHtml: processed.rowHtml,
      filtersData: processed.filtersData,
      showTopFilters: showTopFilters,
      downloadUrl: dlUrl,
      adminCsvIndex: adminCsvIndex,
      truncateAfterAdmin: truncateAfterAdmin,
      removeJustAdmin: removeJustAdmin,
      filterColumnMap: processed.filterColumnMap
    });

    form.clientScriptModulePath = './CL_RR_Tool.js';
    context.response.writePage(form);
  }

  function handlePost(context) {
    var params = context.request.parameters || {};

    var fileId = params.custpage_file_id || '';
    var postedEmp = params.custpage_empid || '';
    var postedTs = params.custpage_ts || '';
    var postedSig = params.custpage_sig || '';

    var authorized = false;
    if (postedEmp && postedTs && postedSig) {
      authorized = verifyUser(postedEmp, postedTs, postedSig);
    }

    if (!authorized) {
      context.response.write(
        '<html><head>' +
        '<script>setTimeout(function(){ window.location.href = ' + JSON.stringify(PORTAL_URL) + '; }, 1200);</script>' +
        '<style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>' +
        '</head><body><div class="message">Session expired. Please log in again.</div></body></html>'
      );
      return;
    }

    var fileObj = file.load({ id: fileId });
    var content = fileObj.getContents();
    var rows = content.split(/\r?\n/);

    var selectedRows = [];
    var selectedJson = params.custpage_selected || '';
    var createdCount = 0;

    if (selectedJson) {
      try {
        var picked = JSON.parse(selectedJson);
        var i;
        for (i = 0; i < picked.length; i++) {
          var p = picked[i] || {};
          var rowId = String(p.rowId || '').trim();
          if (!rowId) continue;

          var csvRow = rows[parseInt(rowId, 10)] || '';
          if (!csvRow) continue;

          var columns = parseCsvLine(csvRow);
          var c;
          for (c = 0; c < columns.length; c++) {
            columns[c] = sanitizeCsvText(unquoteCsv(columns[c]));
          }

          selectedRows.push({
            rowId: rowId,
            qty: safeParseInt(p.qty),
            memo: sanitizeCsvText(p.memo || ''),
            monthStock: safeParseFloat(p.mos),
            columns: columns
          });
        }
      } catch (e) {
        log.error('Bad custpage_selected JSON', e);
      }
    }

    if (!selectedRows.length) {
      var key;
      for (key in params) {
        if (key.indexOf('row_select_') === 0) {
          var fallbackRowId = key.split('_')[2];
          var fallbackQty = safeParseInt(params['qty_input_' + fallbackRowId]);
          var fallbackMemo = sanitizeCsvText(params['memo_input_' + fallbackRowId] || '');
          var fallbackCsvRow = rows[parseInt(fallbackRowId, 10)] || '';
          if (!fallbackCsvRow) continue;

          var fallbackCols = parseCsvLine(fallbackCsvRow);
          var j;
          for (j = 0; j < fallbackCols.length; j++) {
            fallbackCols[j] = sanitizeCsvText(unquoteCsv(fallbackCols[j]));
          }

          selectedRows.push({
            rowId: fallbackRowId,
            qty: fallbackQty,
            memo: fallbackMemo,
            monthStock: 0,
            columns: fallbackCols
          });
        }
      }
    }

    log.debug('selectedRows count', selectedRows.length);

    var k;
    for (k = 0; k < selectedRows.length; k++) {
      var entry = selectedRows[k];
      var cols = entry.columns || [];

      try {
        var monStock = entry.monthStock;
        if (!isFinite(monStock)) monStock = 0;

        record.create({
          type: 'customrecord_mi_planned_po',
          isDynamic: true
        })
        .setValue({ fieldId: 'custrecord_mi_item', value: cols[3] || '' })
        .setValue({ fieldId: 'custrecord_mi_order_qty', value: entry.qty || 0 })
        .setValue({ fieldId: 'custrecord_mi_purchase_memo', value: entry.memo || '' })
        .setValue({ fieldId: 'custrecord_month_of_stocks', value: safeParseFloat(monStock) })
        .setValue({ fieldId: 'custrecord_mi_qty_of_ordered_not_ship', value: safeParseFloat(cols[24]) })
        .setValue({ fieldId: 'custrecord_mi_qty_available', value: safeParseFloat(cols[31]) })
        .setValue({ fieldId: 'custrecord_mi_qty_in_transit', value: safeParseFloat(cols[25]) })
        .setValue({ fieldId: 'custrecord_mi_min_month_qty', value: safeParseFloat(cols[16]) })
        .save();

        createdCount++;
      } catch (e2) {
        log.error('Error creating custom record for row ' + entry.rowId, e2);
      }
    }

    context.response.write(
      '<html><head>' +
      '<meta http-equiv="refresh" content="5;URL=https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2108&deploy=1&compid=4975346&ns-at=AAEJ7tMQmJxVsovhMpsEMUF39xnBuyMwWM4G2T7SnvA62twq8hg&empid=' + encodeURIComponent(postedEmp) + '&ts=' + encodeURIComponent(postedTs) + '&sig=' + encodeURIComponent(postedSig) + '" />' +
      '<style>' +
      'body{font-family:Arial,sans-serif;text-align:center;padding-top:100px;}' +
      '.message-box{display:inline-block;background-color:#f3f9ff;padding:20px 30px;border:1px solid #a3d2f2;border-radius:8px;color:#005b99;font-size:16px;}' +
      '</style></head><body>' +
      '<div class="message-box"><strong>' + createdCount + '</strong> Planned Purchase Order(s) created.<br />You will be redirected to the main page in 5 seconds...</div>' +
      '</body></html>'
    );
  }

  function addHiddenFields(form, selectedEmp, ts, sig) {
    var empField = form.addField({
      id: 'custpage_empid',
      type: ui.FieldType.SELECT,
      label: 'Current Employee',
      source: 'employee'
    });
    empField.defaultValue = selectedEmp || '';
    empField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

    var tsField = form.addField({
      id: 'custpage_ts',
      type: ui.FieldType.TEXT,
      label: 'ts'
    });
    tsField.defaultValue = ts || '';
    tsField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

    var sigField = form.addField({
      id: 'custpage_sig',
      type: ui.FieldType.TEXT,
      label: 'sig'
    });
    sigField.defaultValue = sig || '';
    sigField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });
  }

  function getLatestFilesFromFolders() {
    var result = {
      file1: null,
      file2: null,
      file3: null
    };

    var folderSearchObj = search.create({
      type: 'folder',
      filters: [
        ['internalid', 'anyof', String(SOURCE_FOLDER_1), String(SOURCE_FOLDER_2), String(SOURCE_FOLDER_3)],
        'AND',
        ['file.documentsize', 'greaterthan', '5']
      ],
      columns: [
        search.createColumn({ name: 'internalid', summary: 'GROUP' }),
        search.createColumn({ name: 'internalid', join: 'file', summary: 'MAX' })
      ]
    });

    folderSearchObj.run().each(function (res) {
      var folderId = res.getValue({ name: 'internalid', summary: 'GROUP' });
      var latestFileId = res.getValue({ name: 'internalid', join: 'file', summary: 'MAX' });

      if (String(folderId) === String(SOURCE_FOLDER_1)) result.file1 = latestFileId;
      if (String(folderId) === String(SOURCE_FOLDER_2)) result.file2 = latestFileId;
      if (String(folderId) === String(SOURCE_FOLDER_3)) result.file3 = latestFileId;
      return true;
    });

    return result;
  }

  function buildMergedRows(latestFiles) {
    var rows = [];
    var rows1 = loadRows(latestFiles.file1);
    var rows2 = loadRows(latestFiles.file2);
    var rows3 = loadRows(latestFiles.file3);

    if (rows1.length) {
      rows = rows1.slice();
    }

    if (rows2.length) {
      rows2.shift();
      rows = rows.concat(rows2);
    }

    if (rows3.length) {
      rows3.shift();
      rows = rows.concat(rows3);
    }

    return rows;
  }

  function loadRows(fileId) {
    if (!fileId) return [];
    var fileObj = file.load({ id: fileId });
    var content = fileObj.getContents() || '';
    content = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    return content.split('\n');
  }

  function processRows(opts) {
    var rows = opts.sourceRows;
    var headers = opts.headers;
    var adminCsvIndex = opts.adminCsvIndex;
    var truncateAfterAdmin = opts.truncateAfterAdmin;
    var removeJustAdmin = opts.removeJustAdmin;

    var ITEM_INDEX = 2;
    var VENDOR_INDEX = 35;
    var BRAND_CATEGORY_INDEX = 33;
    var BRAND_INDEX = 32;
    var DEPT_INDEX = 52;
    var POL_INDEX = 36;
    var PRODUCT_INDEX = 53;

    var tableHeaders = getOutputHeaders(headers, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);
    var newContent = [];
    var rowHtml = [];

    var filtersData = {
      items: {},
      vendors: {},
      brands: {},
      brandCategories: {},
      departments: {},
      pols: {},
      products: {}
    };

    var filterColumnMap = {
      item: null,
      vendor: null,
      brand: null,
      brandCategory: null,
      department: null,
      pol: null,
      product: null
    };

    var inventoryMap = getInventoryBalanceMap();
    var alreadyExist = {};

    var finalHeadersForCsv = tableHeaders.slice();
    finalHeadersForCsv.push('Filter');
    newContent.push(finalHeadersForCsv);

    filterColumnMap.item = getOutputIndexForOriginalCsvIndex(ITEM_INDEX, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);
    filterColumnMap.vendor = getOutputIndexForOriginalCsvIndex(VENDOR_INDEX, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);
    filterColumnMap.brand = getOutputIndexForOriginalCsvIndex(BRAND_INDEX, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);
    filterColumnMap.brandCategory = getOutputIndexForOriginalCsvIndex(BRAND_CATEGORY_INDEX, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);
    filterColumnMap.department = getOutputIndexForOriginalCsvIndex(DEPT_INDEX, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);
    filterColumnMap.pol = getOutputIndexForOriginalCsvIndex(POL_INDEX, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);
    filterColumnMap.product = getOutputIndexForOriginalCsvIndex(PRODUCT_INDEX, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);

    var redRows = [];
    var blackRows = [];

    var r;
    for (r = 1; r < rows.length; r++) {
      var rawRow = rows[r];
      if (!rawRow || !rawRow.trim()) continue;

      var columns = parseCsvLine(rawRow);
      if (!columns || !columns.length) continue;

      if (truncateAfterAdmin && adminCsvIndex >= 0) {
        var adminCellVal = sanitizeCsvText(columns[adminCsvIndex] || '').toLowerCase();
        if (adminCellVal !== 'admin portal') {
          continue;
        }
      }

      var calcCols = columns.slice();

      var itemid = sanitizeCsvText(unquoteCsv(calcCols[3] || ''));
      if (!itemid) continue;
      if (alreadyExist[itemid]) continue;
      alreadyExist[itemid] = true;

      addToSetMap(filtersData.items, sanitizeCsvText(unquoteCsv(calcCols[ITEM_INDEX] || '')));
      addToSetMap(filtersData.vendors, sanitizeCsvText(unquoteCsv(calcCols[VENDOR_INDEX] || '')));
      addToSetMap(filtersData.brands, sanitizeCsvText(unquoteCsv(calcCols[BRAND_INDEX] || '')));
      addToSetMap(filtersData.brandCategories, sanitizeCsvText(unquoteCsv(calcCols[BRAND_CATEGORY_INDEX] || '')));
      addToSetMap(filtersData.departments, sanitizeCsvText(unquoteCsv(calcCols[DEPT_INDEX] || '')));
      addToSetMap(filtersData.pols, sanitizeCsvText(unquoteCsv(calcCols[POL_INDEX] || '')));
      addToSetMap(filtersData.products, sanitizeCsvText(unquoteCsv(calcCols[PRODUCT_INDEX] || '')));

      var monthAvg = safeParseFloat(unquoteCsv(calcCols[11]));
      var val43 = safeParseFloat(unquoteCsv(calcCols[25]));
      var val41 = safeParseFloat(unquoteCsv(calcCols[20]));
      var diff = Math.abs(val43 - val41);

      calcCols[25] = String(diff || '');

      var good = 0;
      var bad = 0;
      var hold = 0;
      var inspect = 0;
      var label = 0;
      var total = 0;
      var avail = 0;

      if (inventoryMap[itemid]) {
        good = safeParseFloat(inventoryMap[itemid].good);
        bad = safeParseFloat(inventoryMap[itemid].bad);
        hold = safeParseFloat(inventoryMap[itemid].hold);
        inspect = safeParseFloat(inventoryMap[itemid].inspect);
        label = safeParseFloat(inventoryMap[itemid].label);
        total = safeParseFloat(inventoryMap[itemid].total);
        avail = safeParseFloat(inventoryMap[itemid].avail);
      }

      var col12Val = safeParseFloat(unquoteCsv(calcCols[12]));
      var availToProm = good - col12Val;

      calcCols[12] = numText(availToProm);
      calcCols[13] = numText(good);
      calcCols[14] = numText(bad);
      calcCols[15] = numText(inspect);
      calcCols[16] = numText(label);
      calcCols[17] = numText(hold);
      calcCols[18] = numText(total);
      calcCols[19] = monthAvg ? numText(total / monthAvg) : '0';
      calcCols[23] = numText(total + val41);
      calcCols[24] = monthAvg ? numText((total + val41) / monthAvg) : '0';
      calcCols[26] = numText(total + val43);
      calcCols[27] = monthAvg ? numText((total + val43) / monthAvg) : '0';
      calcCols[31] = numText(avail);
      calcCols[63] = numText(normalizeMovement(monthAvg));
      calcCols[64] = numText(safeParseFloat(normalizeMovement(monthAvg)) * 4);

      var qtytotal = diff + val41 + avail - col12Val;
      var stockingQty = Math.ceil(safeParseFloat(unquoteCsv(calcCols[11])) * 4.5);
      calcCols[11] = numText(normalizeMovement(monthAvg));

      var recommendedQty = 0;
      if (qtytotal < stockingQty) {
        recommendedQty = safeParseFloat((stockingQty - qtytotal).toFixed(2));
      }

      var monthsStock = monthAvg ? ((diff + avail + val41) / monthAvg) : 0;
      var rowColor = 'Black';
      var rowStyle = '';

      if (!isNaN(monthsStock) && monthsStock <= 4.5) {
        rowColor = 'Red';
        rowStyle = ' style="color:red;"';
      }

      var baseCols = calcCols.slice();
      var displayCols = getOutputRowColumns(baseCols, adminCsvIndex, truncateAfterAdmin, removeJustAdmin);
      displayCols.push(rowColor);

      var cleanedCols = [];
      var c;
      for (c = 0; c < displayCols.length; c++) {
        cleanedCols.push(cleanCsvValue(displayCols[c]));
      }
      newContent.push(cleanedCols);

      var rowId = newContent.length - 1;
      var htmlRow = '';
      htmlRow += '<tr' + rowStyle + '>';
      htmlRow += '<td class="sticky-col sticky-1"><input type="checkbox" name="row_select_' + rowId + '" /></td>';
      htmlRow += '<td class="sticky-col sticky-2"><input type="number" name="qty_input_' + rowId + '" min="0" value="' + escapeAttr(String(recommendedQty)) + '" /></td>';
      htmlRow += '<td class="sticky-col sticky-3 month-stock-cell">' + escapeHtml(numText(monthsStock)) + '</td>';
      htmlRow += '<td class="sticky-col sticky-4 cubic-space-cell"></td>';
      htmlRow += '<td class="sticky-col sticky-5 weight-cell"></td>';

      for (c = 0; c < displayCols.length; c++) {
        var cellValue = sanitizeCsvText(unquoteCsv(displayCols[c]));
        if (cellValue === '- None -' || cellValue === 'NaN' || cellValue === 'Infinity') cellValue = '';
        if (cellValue === '.00') cellValue = '0.00';

        if (c === 0) {
          htmlRow += '<td class="sticky-col sticky-6"><input type="text" name="memo_input_' + rowId + '" value="' + escapeAttr(cellValue) + '" /></td>';
        } else {
          htmlRow += '<td>' + escapeHtml(cellValue) + '</td>';
        }
      }

      htmlRow += '</tr>';

      var itemIdNum = safeParseInt(itemid);
      if (rowColor === 'Red') {
        redRows.push({ itemId: itemIdNum, html: htmlRow });
      } else {
        blackRows.push({ itemId: itemIdNum, html: htmlRow });
      }
    }

    redRows.sort(function (a, b) { return a.itemId - b.itemId; });
    blackRows.sort(function (a, b) { return a.itemId - b.itemId; });

    for (r = 0; r < redRows.length; r++) rowHtml.push(redRows[r].html);
    for (r = 0; r < blackRows.length; r++) rowHtml.push(blackRows[r].html);

    return {
      tableHeaders: tableHeaders,
      rowHtml: rowHtml.join(''),
      newContent: newContent,
      filtersData: {
        items: sortedKeys(filtersData.items),
        vendors: sortedKeys(filtersData.vendors),
        brands: sortedKeys(filtersData.brands),
        brandCategories: sortedKeys(filtersData.brandCategories),
        departments: sortedKeys(filtersData.departments),
        pols: sortedKeys(filtersData.pols),
        products: sortedKeys(filtersData.products)
      },
      filterColumnMap: filterColumnMap
    };
  }

  function buildHtml(opts) {
    var tableHeaders = opts.tableHeaders || [];
    var showTopFilters = opts.showTopFilters;
    var filtersData = opts.filtersData || {};
    var downloadUrl = opts.downloadUrl || '';
    var filterColumnMap = opts.filterColumnMap || {};

    var html = '';

    html += '<style>';
    html += 'body{font-family:Arial,sans-serif;}';
    html += '.top-bar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:10px;flex-wrap:wrap;}';
    html += '.left-tools,.right-tools{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}';
    html += '.download-btn{background:#fff;border:1px solid #cbd5e1;border-radius:6px;padding:6px 10px;cursor:pointer;}';
    html += '.download-btn:hover{background:#f8fafc;}';
    html += '.filter-wrap{position:relative;display:inline-block;}';
    html += '.filter-toggle{border:1px solid #cbd5e1;background:#fff;border-radius:6px;padding:6px 10px;cursor:pointer;font-weight:600;}';
    html += '.filter-toggle:hover{background:#f8fafc;}';
    html += '.filter-panel{display:none;position:absolute;top:110%;left:0;background:#fff;border:1px solid #d1d5db;border-radius:8px;padding:12px;width:760px;box-shadow:0 8px 24px rgba(0,0,0,.12);z-index:5000;}';
    html += '.filter-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;}';
    html += '.filter-grid label{display:block;font-size:12px;font-weight:700;margin-bottom:4px;}';
    html += '.filter-grid select{width:100%;min-height:110px;}';
    html += '.totals-bar{display:inline-flex;align-items:center;gap:8px;padding:6px 12px;border:1px solid #a3d2f2;border-radius:999px;background:#f3f9ff;}';
    html += '.totals-label{font-size:12px;color:#005b99;font-weight:700;text-transform:uppercase;}';
    html += '.totals-value{font-weight:700;font-size:14px;color:#003f6b;}';
    html += '.table-shell{border:1px solid #d1d5db;border-radius:8px;overflow:hidden;}';
    html += '.top-scroll{height:16px;overflow-x:auto;overflow-y:hidden;border-bottom:1px solid #d1d5db;background:#f8fafc;}';
    html += '.top-scroll-inner{height:1px;}';
    html += '.table-container{max-height:850px;overflow:auto;position:relative;}';
    html += '#excelTable{border-collapse:separate;border-spacing:0;table-layout:auto;min-width:100%;width:max-content;background:#fff;}';
    html += '#excelTable th,#excelTable td{border-right:1px solid #d1d5db;border-bottom:1px solid #d1d5db;padding:8px 10px;white-space:nowrap;font-size:12px;background:#fff;}';
    html += '#excelTable thead th{position:sticky;top:0;background:#f3f4f6;z-index:30;}';
    html += '.sticky-col{position:sticky;background:#f9fafb !important;z-index:20;}';
    html += '#excelTable thead .sticky-col{z-index:40;}';
    html += '.sticky-1{left:0;min-width:60px;max-width:60px;width:60px;}';
    html += '.sticky-2{left:60px;min-width:100px;max-width:100px;width:100px;}';
    html += '.sticky-3{left:160px;min-width:110px;max-width:110px;width:110px;}';
    html += '.sticky-4{left:270px;min-width:110px;max-width:110px;width:110px;}';
    html += '.sticky-5{left:380px;min-width:90px;max-width:90px;width:90px;}';
    html += '.sticky-6{left:470px;min-width:220px;max-width:220px;width:220px;}';
    html += 'input[type="number"],input[type="text"]{width:100%;box-sizing:border-box;padding:4px 6px;}';
    html += '.th-filter-wrap{display:inline-flex;align-items:center;gap:6px;}';
    html += '.th-filter-btn{cursor:pointer;border:1px solid #cbd5e1;background:#fff;padding:2px 5px;border-radius:4px;font-size:11px;}';
    html += '.th-filter-btn:hover{background:#f8fafc;}';
    html += '.th-filter-panel{position:fixed;top:0;left:0;width:260px;max-height:320px;overflow:auto;background:#fff;border:1px solid #cbd5e1;border-radius:8px;padding:10px;box-shadow:0 8px 24px rgba(0,0,0,.12);z-index:99999;}';
    html += '.th-filter-panel .hdr{font-weight:700;font-size:12px;margin-bottom:6px;}';
    html += '.th-filter-panel input[type="search"]{width:100%;box-sizing:border-box;padding:6px 8px;margin-bottom:8px;}';
    html += '.th-filter-panel .list{max-height:190px;overflow:auto;border:1px solid #e5e7eb;border-radius:6px;padding:6px;}';
    html += '.th-filter-panel .row{display:flex;align-items:center;gap:8px;padding:2px 0;font-size:12px;}';
    html += '.th-filter-panel .actions{display:flex;justify-content:space-between;gap:8px;margin-top:8px;}';
    html += '.th-filter-panel .actions button{padding:3px 7px;font-size:11px;border-radius:4px;border:1px solid #cbd5e1;background:#fff;cursor:pointer;}';
    html += '.th-filter-active .th-filter-btn{border-color:#2563eb;background:#eff6ff;}';
    html += '</style>';

    html += '<div class="top-bar">';
    html += '<div class="left-tools">';
    html += '<button id="downloadCsvBtn" class="download-btn" type="button">Download CSV</button>';

    if (showTopFilters) {
      html += '<div class="filter-wrap" id="filterWrapper">';
      html += '<button type="button" class="filter-toggle" id="filterToggle">Filters</button>';
      html += '<div class="filter-panel" id="filterPanel">';
      html += '<div class="filter-grid">';
      html += buildTopSelect('itemFilter', 'Item', filtersData.items);
      html += buildTopSelect('vendorFilter', 'Vendor', filtersData.vendors);
      html += buildTopSelect('brandFilter', 'Brand', filtersData.brands);
      html += buildTopSelect('brandCatFilter', 'Brand Category', filtersData.brandCategories);
      html += buildTopSelect('deptFilter', 'Department', filtersData.departments);
      html += buildTopSelect('polFilter', 'P.O.L', filtersData.pols);
      html += buildTopSelect('productFilter', 'Product Type', filtersData.products);
      html += '</div></div></div>';
    }

    html += '</div>';
    html += '<div class="right-tools">';
    html += '<div class="totals-bar"><span class="totals-label">Total Cubic Space</span><span id="totalCubicValue" class="totals-value">0</span></div>';
    html += '<div class="totals-bar"><span class="totals-label">Total Weight</span><span id="totalWeightValue" class="totals-value">0</span></div>';
    html += '</div>';
    html += '</div>';

    html += '<div class="table-shell">';
    html += '<div id="topScroll" class="top-scroll"><div id="topScrollInner" class="top-scroll-inner"></div></div>';
    html += '<div class="table-container" id="tableContainer">';
    html += '<table id="excelTable">';
    html += '<thead><tr>';
    html += '<th class="sticky-col sticky-1">Select</th>';
    html += '<th class="sticky-col sticky-2">Order Qty</th>';
    html += '<th class="sticky-col sticky-3">Month of Stock</th>';
    html += '<th class="sticky-col sticky-4">Item Space</th>';
    html += '<th class="sticky-col sticky-5">Weight</th>';

    var h;
    for (h = 0; h < tableHeaders.length; h++) {
      var domIndex = h + 5;
      var stickyClass = (h === 0 ? ' sticky-col sticky-6' : '');
      html += '<th class="' + stickyClass + '" data-col-index="' + domIndex + '">' +
        '<span class="th-filter-wrap">' +
        '<span>' + escapeHtml(sanitizeCsvText(unquoteCsv(tableHeaders[h]))) + '</span>' +
        '<button type="button" class="th-filter-btn" data-col="' + domIndex + '">▾</button>' +
        '</span></th>';
    }
    html += '</tr></thead><tbody>';
    html += opts.rowHtml || '';
    html += '</tbody></table></div></div>';
    html += '<div id="filter-portal"></div>';

    html += '<script>';
    html += '(function(){';
    html += 'var DOWNLOAD_URL=' + JSON.stringify(downloadUrl) + ';';
    html += 'var FILTER_COL_MAP=' + JSON.stringify(filterColumnMap) + ';';
    html += 'var topScroll=document.getElementById("topScroll");';
    html += 'var topScrollInner=document.getElementById("topScrollInner");';
    html += 'var container=document.getElementById("tableContainer");';
    html += 'var table=document.getElementById("excelTable");';
    html += 'var portal=document.getElementById("filter-portal");';
    html += 'var activeHeaderFilters={};';

    html += 'function syncTopScrollWidth(){if(!table||!topScrollInner)return;topScrollInner.style.width=Math.max(table.scrollWidth,container.clientWidth+2)+"px";}';
    html += 'if(topScroll&&container){topScroll.addEventListener("scroll",function(){container.scrollLeft=topScroll.scrollLeft;});container.addEventListener("scroll",function(){topScroll.scrollLeft=container.scrollLeft;});}';
    html += 'window.addEventListener("resize",syncTopScrollWidth);setTimeout(syncTopScrollWidth,50);setTimeout(syncTopScrollWidth,300);';

    html += 'var downloadBtn=document.getElementById("downloadCsvBtn");';
    html += 'if(downloadBtn){downloadBtn.addEventListener("click",function(e){e.preventDefault();if(DOWNLOAD_URL)window.open(DOWNLOAD_URL,"_blank");});}';

    html += 'var filterWrapper=document.getElementById("filterWrapper");';
    html += 'var filterToggle=document.getElementById("filterToggle");';
    html += 'var filterPanel=document.getElementById("filterPanel");';
    html += 'if(filterWrapper&&filterToggle&&filterPanel){';
    html += 'filterToggle.addEventListener("click",function(e){e.stopPropagation();filterPanel.style.display=(filterPanel.style.display==="block"?"none":"block");});';
    html += 'document.addEventListener("click",function(e){if(!filterWrapper.contains(e.target))filterPanel.style.display="none";});';
    html += '}';

    html += 'function getSelectedValues(selectId){var el=document.getElementById(selectId);if(!el)return [];var out=[];for(var i=0;i<el.options.length;i++){if(el.options[i].selected)out.push(String(el.options[i].value||"").toLowerCase());}return out;}';
    html += 'function textOfCell(row, idx){if(idx==null||idx<0)return "";var cell=row.cells[idx];if(!cell)return "";return String(cell.textContent||"").trim().toLowerCase();}';

    html += 'function passesTopFilters(row){';
    html += 'var selectedItems=getSelectedValues("itemFilter");';
    html += 'var selectedVendors=getSelectedValues("vendorFilter");';
    html += 'var selectedBrands=getSelectedValues("brandFilter");';
    html += 'var selectedBrandCats=getSelectedValues("brandCatFilter");';
    html += 'var selectedDepts=getSelectedValues("deptFilter");';
    html += 'var selectedPols=getSelectedValues("polFilter");';
    html += 'var selectedProducts=getSelectedValues("productFilter");';

    html += 'var item=textOfCell(row, FILTER_COL_MAP.item != null ? FILTER_COL_MAP.item + 5 : -1);';
    html += 'var vendor=textOfCell(row, FILTER_COL_MAP.vendor != null ? FILTER_COL_MAP.vendor + 5 : -1);';
    html += 'var brand=textOfCell(row, FILTER_COL_MAP.brand != null ? FILTER_COL_MAP.brand + 5 : -1);';
    html += 'var brandCat=textOfCell(row, FILTER_COL_MAP.brandCategory != null ? FILTER_COL_MAP.brandCategory + 5 : -1);';
    html += 'var dept=textOfCell(row, FILTER_COL_MAP.department != null ? FILTER_COL_MAP.department + 5 : -1);';
    html += 'var pol=textOfCell(row, FILTER_COL_MAP.pol != null ? FILTER_COL_MAP.pol + 5 : -1);';
    html += 'var product=textOfCell(row, FILTER_COL_MAP.product != null ? FILTER_COL_MAP.product + 5 : -1);';

    html += 'var ok=true;';
    html += 'if(selectedItems.length && selectedItems.indexOf(item)===-1) ok=false;';
    html += 'if(selectedVendors.length && selectedVendors.indexOf(vendor)===-1) ok=false;';
    html += 'if(selectedBrands.length && selectedBrands.indexOf(brand)===-1) ok=false;';
    html += 'if(selectedBrandCats.length && selectedBrandCats.indexOf(brandCat)===-1) ok=false;';
    html += 'if(selectedDepts.length && selectedDepts.indexOf(dept)===-1) ok=false;';
    html += 'if(selectedPols.length && selectedPols.indexOf(pol)===-1) ok=false;';
    html += 'if(selectedProducts.length && selectedProducts.indexOf(product)===-1) ok=false;';
    html += 'return ok;';
    html += '}';

    html += 'function normalizeVal(v){v=String(v==null?"":v).trim();if(/^[-\\s]*none[-\\s]*$/i.test(v))v="";if(/^nan$/i.test(v))v="";return v.toLowerCase();}';
    html += 'function getBodyRows(){return table && table.tBodies && table.tBodies[0] ? Array.prototype.slice.call(table.tBodies[0].rows) : [];}';
    html += 'function rowPassesHeaderFilters(row){for(var k in activeHeaderFilters){if(!activeHeaderFilters.hasOwnProperty(k))continue;var set=activeHeaderFilters[k];if(!set||!set.size)continue;var val=normalizeVal(row.cells[parseInt(k,10)] ? row.cells[parseInt(k,10)].textContent : "");if(!set.has(val))return false;}return true;}';
    html += 'function applyAllFilters(){var rows=getBodyRows();for(var i=0;i<rows.length;i++){var row=rows[i];var show=passesTopFilters(row)&&rowPassesHeaderFilters(row);row.style.display=show?"":"none";}updateTotals();}';

    html += '["itemFilter","vendorFilter","brandFilter","brandCatFilter","deptFilter","polFilter","productFilter"].forEach(function(id){var el=document.getElementById(id);if(el)el.addEventListener("change",applyAllFilters);});';

    html += 'function getCountsForColumn(colIdx){var rows=getBodyRows();var counts={};for(var i=0;i<rows.length;i++){var row=rows[i];if(!passesTopFilters(row))continue;var ok=true;for(var k in activeHeaderFilters){if(!activeHeaderFilters.hasOwnProperty(k))continue;if(parseInt(k,10)===colIdx)continue;var set=activeHeaderFilters[k];if(!set||!set.size)continue;var vv=normalizeVal(row.cells[parseInt(k,10)] ? row.cells[parseInt(k,10)].textContent : "");if(!set.has(vv)){ok=false;break;}}if(!ok)continue;var val=normalizeVal(row.cells[colIdx] ? row.cells[colIdx].textContent : "");counts[val]=(counts[val]||0)+1;}return counts;}';

    html += 'function allValuesForColumn(colIdx){var rows=getBodyRows();var map={};for(var i=0;i<rows.length;i++){var label=String(rows[i].cells[colIdx] ? rows[i].cells[colIdx].textContent : "").trim();if(label===".00")label="0.00";if(label==="- None -"||label==="NaN"||label==="Infinity")label="";var key=normalizeVal(label);if(!map.hasOwnProperty(key))map[key]=key===""?"(blank)":label;}var keys=Object.keys(map).sort(function(a,b){if(a==="")return 1;if(b==="")return -1;return String(map[a]).localeCompare(String(map[b]),undefined,{numeric:true,sensitivity:"base"});});return {keys:keys,map:map};}';

    html += 'function closeHeaderPanels(){var open=portal.querySelector(".th-filter-panel");if(open)open.remove();var activeThs=table.querySelectorAll("th.th-filter-active");for(var i=0;i<activeThs.length;i++)activeThs[i].classList.remove("th-filter-active");}';

    html += 'function openHeaderFilter(btn,colIdx){closeHeaderPanels();var th=btn.closest("th");if(!th)return;';
    html += 'var grouped=allValuesForColumn(colIdx);var counts=getCountsForColumn(colIdx);var existing=activeHeaderFilters[colIdx]||null;var temp=new Set();';
    html += 'if(existing&&existing.size){grouped.keys.forEach(function(k){if(existing.has(k))temp.add(k);});}else{grouped.keys.forEach(function(k){if((counts[k]||0)>0)temp.add(k);});}';
    html += 'var panel=document.createElement("div");panel.className="th-filter-panel";';
    html += 'panel.innerHTML=\'<div class="hdr">Filter</div><input type="search" placeholder="Search values..."><div class="list"></div><div class="actions"><button type="button" data-act="clear">Clear</button><div style="display:flex;gap:6px;"><button type="button" data-act="selectall">Select all</button><button type="button" data-act="deselectall">Deselect all</button><button type="button" data-act="apply">Apply</button></div></div>\';';
    html += 'var list=panel.querySelector(".list");';
    html += 'function render(ft){list.innerHTML="";ft=String(ft||"").toLowerCase();grouped.keys.forEach(function(k){var label=grouped.map[k]||"(blank)";if(ft&&label.toLowerCase().indexOf(ft)===-1)return;var row=document.createElement("div");row.className="row";row.dataset.key=k;var id="f_"+colIdx+"_"+Math.random().toString(36).slice(2);row.innerHTML=\'<input type="checkbox" id="\'+id+\'" \'+(temp.has(k)?"checked":"")+\'><label for="\'+id+\'">\'+label+((counts[k]||0)?(" ("+counts[k]+")"):"")+\'</label>\';list.appendChild(row);});}';
    html += 'render("");';
    html += 'panel.querySelector("input[type=search]").addEventListener("input",function(){render(this.value);});';
    html += 'list.addEventListener("change",function(e){var cb=e.target;if(!cb||cb.type!=="checkbox")return;var row=cb.closest(".row");if(!row)return;var key=row.dataset.key||"";if(cb.checked)temp.add(key);else temp.delete(key);});';
    html += 'panel.querySelector(".actions").addEventListener("click",function(e){var act=e.target.getAttribute("data-act");if(!act)return;if(act==="clear"){delete activeHeaderFilters[colIdx];closeHeaderPanels();applyAllFilters();return;}if(act==="selectall"){var cbs=list.querySelectorAll("input[type=checkbox]");for(var i=0;i<cbs.length;i++){cbs[i].checked=true;temp.add(cbs[i].closest(".row").dataset.key||"");}return;}if(act==="deselectall"){var cbs2=list.querySelectorAll("input[type=checkbox]");for(var j=0;j<cbs2.length;j++){cbs2[j].checked=false;temp.delete(cbs2[j].closest(".row").dataset.key||"");}return;}if(act==="apply"){if(temp.size===grouped.keys.length){delete activeHeaderFilters[colIdx];}else{activeHeaderFilters[colIdx]=new Set(temp);}closeHeaderPanels();if(activeHeaderFilters[colIdx]&&activeHeaderFilters[colIdx].size){th.classList.add("th-filter-active");}else{th.classList.remove("th-filter-active");}applyAllFilters();}});';
    html += 'portal.appendChild(panel);th.classList.add("th-filter-active");';
    html += 'var rect=btn.getBoundingClientRect();var pw=260;var ph=panel.offsetHeight||280;var left=Math.max(8,Math.min((window.innerWidth-pw-8),(rect.right-pw)));var top=(rect.bottom+ph+8<window.innerHeight)?(rect.bottom+6):Math.max(8,rect.top-ph-6);panel.style.left=Math.round(left)+"px";panel.style.top=Math.round(top)+"px";';
    html += '}';

    html += 'table.tHead.addEventListener("click",function(e){var btn=e.target.closest(".th-filter-btn");if(!btn)return;e.stopPropagation();openHeaderFilter(btn,parseInt(btn.getAttribute("data-col"),10));});';
    html += 'document.addEventListener("click",function(e){if(!portal.contains(e.target)&&!e.target.closest(".th-filter-btn"))closeHeaderPanels();});';

    html += 'function parseNum(v){v=String(v||"").replace(/,/g,"").trim();var n=parseFloat(v);return isNaN(n)?0:n;}';
    html += 'function updateTotals(){var rows=getBodyRows();var totalCubic=0,totalWeight=0;for(var i=0;i<rows.length;i++){if(rows[i].style.display==="none")continue;var cubicCell=rows[i].querySelector(".cubic-space-cell");var weightCell=rows[i].querySelector(".weight-cell");if(cubicCell)totalCubic+=parseNum(cubicCell.textContent);if(weightCell)totalWeight+=parseNum(weightCell.textContent);}var cubicEl=document.getElementById("totalCubicValue");var weightEl=document.getElementById("totalWeightValue");if(cubicEl)cubicEl.textContent=totalCubic.toFixed(2);if(weightEl)weightEl.textContent=totalWeight.toFixed(2);}';

    html += 'var form=document.querySelector("form");var hiddenSelected=document.getElementById("custpage_selected");';
    html += 'function collectSelectedRows(){if(!hiddenSelected)return;var out=[];var rows=getBodyRows();for(var i=0;i<rows.length;i++){var tr=rows[i];var cb=tr.querySelector(\'input[type="checkbox"][name^="row_select_"]\');if(cb&&cb.checked){var rowId=(cb.name||"").split("_").pop();var qtyInput=tr.querySelector(\'input[type="number"][name^="qty_input_"]\');var memoInput=tr.querySelector(\'input[type="text"][name^="memo_input_"]\');var mosCell=tr.querySelector(".month-stock-cell");out.push({rowId:rowId,qty:qtyInput?qtyInput.value:"0",memo:memoInput?memoInput.value:"",mos:mosCell?mosCell.textContent:"0"});}}hiddenSelected.value=JSON.stringify(out);}';
    html += 'if(form){form.addEventListener("submit",function(){collectSelectedRows();closeHeaderPanels();});}';

    html += 'applyAllFilters();syncTopScrollWidth();';
    html += '})();';
    html += '</script>';

    return html;
  }

  function getOutputHeaders(headers, adminCsvIndex, truncateAfterAdmin, removeJustAdmin) {
    var out = [];
    var i;

    if (adminCsvIndex >= 0) {
      if (truncateAfterAdmin) {
        for (i = 0; i < adminCsvIndex; i++) {
          out.push(sanitizeCsvText(unquoteCsv(headers[i] || '')));
        }
        return out;
      }

      if (removeJustAdmin) {
        for (i = 0; i < headers.length; i++) {
          if (i !== adminCsvIndex) out.push(sanitizeCsvText(unquoteCsv(headers[i] || '')));
        }
        return out;
      }
    }

    for (i = 0; i < headers.length; i++) {
      out.push(sanitizeCsvText(unquoteCsv(headers[i] || '')));
    }
    return out;
  }

  function getOutputRowColumns(baseCols, adminCsvIndex, truncateAfterAdmin, removeJustAdmin) {
    var out = [];
    var i;

    if (adminCsvIndex >= 0) {
      if (truncateAfterAdmin) {
        for (i = 0; i < adminCsvIndex; i++) out.push(baseCols[i]);
        return out;
      }

      if (removeJustAdmin) {
        for (i = 0; i < baseCols.length; i++) {
          if (i !== adminCsvIndex) out.push(baseCols[i]);
        }
        return out;
      }
    }

    return baseCols.slice();
  }

  function getOutputIndexForOriginalCsvIndex(originalIndex, adminCsvIndex, truncateAfterAdmin, removeJustAdmin) {
    if (truncateAfterAdmin && adminCsvIndex >= 0) {
      if (originalIndex >= adminCsvIndex) return null;
      return originalIndex;
    }

    if (removeJustAdmin && adminCsvIndex >= 0) {
      if (originalIndex === adminCsvIndex) return null;
      if (originalIndex > adminCsvIndex) return originalIndex - 1;
      return originalIndex;
    }

    return originalIndex;
  }

  function buildTopSelect(id, label, values) {
    var html = '<div><label for="' + escapeAttr(id) + '">' + escapeHtml(label) + '</label><select id="' + escapeAttr(id) + '" multiple size="6">';
    var i;
    for (i = 0; i < values.length; i++) {
      html += '<option value="' + escapeAttr(values[i]) + '">' + escapeHtml(values[i]) + '</option>';
    }
    html += '</select></div>';
    return html;
  }

  function buildCsvString(rows) {
    var out = [];
    var i, j;
    for (i = 0; i < rows.length; i++) {
      var line = [];
      for (j = 0; j < rows[i].length; j++) {
        line.push(cleanCsvValue(rows[i][j]));
      }
      out.push(line.join(','));
    }
    return out.join('\n');
  }

  function saveCsvFile(name, contents, folderId, isOnline) {
    var f = file.create({
      name: name,
      fileType: file.Type.CSV,
      contents: contents,
      encoding: file.Encoding.UTF8,
      folder: folderId,
      isOnline: !!isOnline
    });
    return f.save();
  }

  function getInventoryBalanceMap() {
    var resultMap = {};

    var inventorybalanceSearchObj = search.create({
      type: 'inventorybalance',
      filters: [
        ['status', 'anyof', '6', '1', '3', '5', '8'],
        'AND',
        ['item.isinactive', 'is', 'F']
      ],
      columns: [
        search.createColumn({ name: 'item', summary: 'GROUP' }),
        search.createColumn({ name: 'formulanumeric', summary: 'SUM', formula: "case when {status} = 'Good' then {onhand} else 0 end" }),
        search.createColumn({ name: 'formulanumeric1', summary: 'SUM', formula: "case when {status} = 'Deviation' then {onhand} else 0 end" }),
        search.createColumn({ name: 'formulanumeric2', summary: 'SUM', formula: "case when {status} = 'Hold' then {onhand} else 0 end" }),
        search.createColumn({ name: 'formulanumeric3', summary: 'SUM', formula: "case when {status} = 'Inspection' then {onhand} else 0 end" }),
        search.createColumn({ name: 'formulanumeric4', summary: 'SUM', formula: "case when {status} = 'Label' then {onhand} else 0 end" }),
        search.createColumn({ name: 'available', summary: 'SUM' })
      ]
    });

    inventorybalanceSearchObj.run().each(function (res) {
      var itemId = res.getValue({ name: 'item', summary: 'GROUP' });
      var goodQty = safeParseFloat(res.getValue({ name: 'formulanumeric', summary: 'SUM' }));
      var badQty = safeParseFloat(res.getValue({ name: 'formulanumeric1', summary: 'SUM' }));
      var holdQty = safeParseFloat(res.getValue({ name: 'formulanumeric2', summary: 'SUM' }));
      var inspectQty = safeParseFloat(res.getValue({ name: 'formulanumeric3', summary: 'SUM' }));
      var labelQty = safeParseFloat(res.getValue({ name: 'formulanumeric4', summary: 'SUM' }));
      var avail = safeParseFloat(res.getValue({ name: 'available', summary: 'SUM' }));
      var total = safeParseFloat((goodQty + badQty).toFixed(2));

      resultMap[String(itemId)] = {
        good: goodQty,
        bad: badQty,
        hold: holdQty,
        inspect: inspectQty,
        label: labelQty,
        total: total,
        avail: avail
      };
      return true;
    });

    return resultMap;
  }

  function signCron(ts) {
    var secret = getSecret();
    var h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
    h.update({ input: 'CRON|' + ts + '|' + secret });
    return h.digest({ outputEncoding: crypto.Encoding.HEX });
  }

  function verifyCron(ts, sig) {
    if (!ts || !sig) return false;
    if (Math.abs(Date.now() - safeParseInt(ts)) > TOKEN_TTL_MS) return false;
    try {
      return signCron(ts) === sig;
    } catch (e) {
      log.error('verifyCron token', e);
      return false;
    }
  }

  function signUser(empid, ts) {
    var secret = getSecret();
    var h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
    h.update({ input: String(empid) + '|' + String(ts) + '|' + secret });
    return h.digest({ outputEncoding: crypto.Encoding.HEX });
  }

  function verifyUser(empid, ts, sig) {
    if (!empid || !ts || !sig) return false;
    if (Math.abs(Date.now() - safeParseInt(ts)) > TOKEN_TTL_MS) return false;
    try {
      return signUser(empid, ts) === sig;
    } catch (e) {
      log.error('verify token', e);
      return false;
    }
  }

  function getSecret() {
    return runtime.getCurrentScript().getParameter({ name: 'custscript_portal_secret' }) || 'change-me';
  }

  function parseCsvLine(line) {
    if (line == null) return [];
    var result = [];
    var current = '';
    var inQuotes = false;
    var i;
    var ch;
    var next;

    line = String(line);

    for (i = 0; i < line.length; i++) {
      ch = line.charAt(i);

      if (ch === '"') {
        next = line.charAt(i + 1);
        if (inQuotes && next === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === ',' && !inQuotes) {
        result.push(current);
        current = '';
      } else {
        current += ch;
      }
    }

    result.push(current);
    return result;
  }

  function unquoteCsv(v) {
    v = String(v == null ? '' : v);
    if (v.length >= 2 && v.charAt(0) === '"' && v.charAt(v.length - 1) === '"') {
      v = v.substring(1, v.length - 1).replace(/""/g, '"');
    }
    return v;
  }

  function sanitizeCsvText(value) {
    if (value == null) return '';
    var s = String(value);

    s = s.replace(/\uFEFF/g, '');
    s = s.replace(/[\u200B-\u200D\u2060]/g, '');
    s = s.replace(/\u00A0/g, ' ');
    s = s.replace(/[“”]/g, '"');
    s = s.replace(/[‘’]/g, "'");
    s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
    s = s.replace(/\r/g, ' ').replace(/\n/g, ' ');

    return s.trim();
  }

  function cleanCsvValue(value) {
    var v = sanitizeCsvText(unquoteCsv(value));

    if (v === '- None -' || v === 'NaN' || v === 'Infinity') return '';
    if (v === '.00') v = '0.00';

    if (/[",\r\n]/.test(v)) {
      v = '"' + v.replace(/"/g, '""') + '"';
    }

    return v;
  }

  function normHeader(h) {
    return String(h || '')
      .replace(/^[\uFEFF\s"]+|[\s"]+$/g, '')
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  function normalizeMovement(val) {
    if (val === null || val === undefined) return 'No Movement';
    var s = String(val).replace(/,/g, '').trim();
    if (!s) return 'No Movement';
    var n = parseFloat(s);
    if (!isFinite(n) || n <= 0) return 'No Movement';
    return n.toFixed(2);
  }

  function numText(v) {
    if (!isFinite(safeParseFloat(v))) return '0';
    return String(safeParseFloat(v));
  }

  function safeParseFloat(val) {
    var n = parseFloat(String(val == null ? '' : val).replace(/,/g, '').trim());
    return isNaN(n) ? 0 : n;
  }

  function safeParseInt(val) {
    var n = parseInt(String(val == null ? '' : val).replace(/,/g, '').trim(), 10);
    return isNaN(n) ? 0 : n;
  }

  function addToSetMap(mapObj, value) {
    value = sanitizeCsvText(value);
    if (!value) return;
    mapObj[value] = true;
  }

  function sortedKeys(obj) {
    var arr = [];
    var k;
    for (k in obj) {
      if (obj.hasOwnProperty(k)) arr.push(k);
    }
    arr.sort(function (a, b) {
      return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
    });
    return arr;
  }

  function escapeHtml(str) {
    str = String(str == null ? '' : str);
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function escapeAttr(str) {
    return escapeHtml(str);
  }

  return {
    onRequest: onRequest
  };
});