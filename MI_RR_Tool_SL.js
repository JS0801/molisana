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
  var RETURN_URL_BASE = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2108&deploy=1&compid=4975346&ns-at=AAEJ7tMQmJxVsovhMpsEMUF39xnBuyMwWM4G2T7SnvA62twq8hg';
  var SECRET = runtime.getCurrentScript().getParameter({ name: 'custscript_portal_secret' }) || 'change-me';
  var TOKEN_TTL_MS = 30 * 60 * 1000;

  var SOURCE_FOLDERS = {
    MAIN: 402335,
    SECONDARY: 402334,
    THIRD: 413248
  };

  var DOWNLOAD_FOLDER_UI = 378271;
  var DOWNLOAD_FOLDER_CRON = 279208;

  var HEADER_INDEX = {
    ITEM: 2,
    VENDOR: 35,
    BRAND_CATEGORY: 33,
    BRAND: 32,
    DEPT: 52,
    POL: 36,
    PRODUCT: 53
  };

  function onRequest(context) {
    try {
      if (context.request.method === 'GET') {
        handleGet(context);
      } else {
        handlePost(context);
      }
    } catch (e) {
      log.error('onRequest error', e);
      context.response.write(
        '<html><body style="font-family:Arial;padding:20px;color:red;">Error: ' +
        escapeHtml(e && e.message ? e.message : String(e)) +
        '</body></html>'
      );
    }
  }

  function handleGet(context) {
    var q = context.request.parameters || {};
    var typeParam = String(q.type || '3');
    var showTopFilters = typeParam !== '4';
    var formName = 'MI Reorder Tool (' + (typeParam === '4' ? 'Basic' : 'Admin') + ')';

    var mode = String(q.mode || '').toLowerCase();
    var cronTs = q.cronts || '';
    var cronSig = q.cronsig || '';
    var isCron = mode === 'cron' && verifyCron(cronTs, cronSig);

    var empid = q.empid || '';
    var ts = q.ts || '';
    var sig = q.sig || '';
    var selectedEmp = q.custpage_id || '';

    if (empid && ts && sig && verify(empid, ts, sig)) {
      selectedEmp = empid;
    }

    if (!selectedEmp && !isCron) {
      writeLoginRequired(context.response);
      return;
    }

    var form = ui.createForm({ title: formName });

    var empField = form.addField({
      id: 'custpage_empid',
      type: ui.FieldType.SELECT,
      label: 'Current Employee',
      source: 'employee'
    });
    empField.defaultValue = selectedEmp;
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

    var fileField = form.addField({
      id: 'custpage_file_id',
      type: ui.FieldType.TEXT,
      label: 'File ID'
    });
    fileField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

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

    var latestFiles = getLatestFiles();
    if (!latestFiles.mainId) {
      htmlField.defaultValue = '<p style="color:red;">No file found in the specified folder.</p>';
      context.response.writePage(form);
      return;
    }

    var mainFile = file.load({ id: latestFiles.mainId });
    var secondaryFile = latestFiles.secondaryId ? file.load({ id: latestFiles.secondaryId }) : null;
    var thirdFile = latestFiles.thirdId ? file.load({ id: latestFiles.thirdId }) : null;

    var mainRows = splitCsvLines(mainFile.getContents());
    var secondRows = secondaryFile ? splitCsvLines(secondaryFile.getContents()) : [];
    var thirdRows = thirdFile ? splitCsvLines(thirdFile.getContents()) : [];

    if (!mainRows.length) {
      htmlField.defaultValue = '<p style="color:red;">Source file is empty.</p>';
      context.response.writePage(form);
      return;
    }

    if (secondRows.length) secondRows.shift();
    if (thirdRows.length) thirdRows.shift();

    var rows = mainRows.concat(secondRows).concat(thirdRows);
    var headers = parseCsvLine(rows[0]);

    var headerMeta = buildHeaderMeta(headers, typeParam);
    var balances = getInventoryBalanceMap();

    var buildResult = buildOutputRows(rows, headerMeta, balances);
    var cleanedCsvContent = buildCsvContent(buildResult.outputRows);

    var uiFile = file.create({
      name: 'Download_' + sanitizeFileName(mainFile.name),
      fileType: file.Type.CSV,
      contents: cleanedCsvContent,
      encoding: file.Encoding.UTF8,
      folder: DOWNLOAD_FOLDER_UI,
      isOnline: true
    });
    var newFileId = uiFile.save();
    fileField.defaultValue = String(newFileId);

    var cronFile = file.create({
      name: 'RR Tool Details.csv',
      fileType: file.Type.CSV,
      contents: cleanedCsvContent,
      encoding: file.Encoding.UTF8,
      folder: DOWNLOAD_FOLDER_CRON,
      isOnline: true
    });
    var newFileId2 = cronFile.save();
    var reloadFile = file.load({ id: newFileId2 });
    var dlUrl = reloadFile.url || '';

    if (isCron) {
      context.response.setHeader({ name: 'Content-Type', value: 'application/json' });
      context.response.write(JSON.stringify({
        ok: true,
        fileId: newFileId,
        url: dlUrl,
        name: reloadFile.name
      }));
      return;
    }

    htmlField.defaultValue = buildHtml({
      downloadUrl: dlUrl,
      typeParam: typeParam,
      showTopFilters: showTopFilters,
      headersForOutput: headerMeta.headersForOutput,
      rowsHtml: buildResult.rowsHtml,
      filterSets: buildResult.filterSets,
      adminCsvIndex: headerMeta.adminCsvIndex,
      truncateAfterAdmin: headerMeta.truncateAfterAdmin,
      removeJustAdmin: headerMeta.removeJustAdmin
    });

    form.clientScriptModulePath = './CL_RR_Tool.js';
    context.response.writePage(form);
  }

  function handlePost(context) {
    var params = context.request.parameters || {};
    var fileId = params.custpage_file_id;

    var postedEmp = params.custpage_empid || '';
    var postedTs = params.custpage_ts || '';
    var postedSig = params.custpage_sig || '';

    var authorized = false;
    if (postedEmp && postedTs && postedSig) {
      authorized = verify(postedEmp, postedTs, postedSig);
    }

    if (!authorized) {
      writeSessionExpired(context.response);
      return;
    }

    var fileObj = file.load({ id: fileId });
    var content = fileObj.getContents();
    var rows = splitCsvLines(content);

    var selectedRows = [];
    var createdCount = 0;

    if (params.custpage_selected) {
      try {
        var picked = JSON.parse(params.custpage_selected);
        picked.forEach(function (p) {
          var rowId = String(p.rowId || '').trim();
          var qty = parseInt(p.qty, 10) || 0;
          var memo = p.memo || '';
          var mos = parseFloat(p.mos);
          if (!rowId || !rows[rowId]) return;

          var columns = parseCsvLine(rows[rowId]).map(function (val) {
            return cleanDisplayValue(val);
          });

          selectedRows.push({
            rowId: rowId,
            qty: qty,
            memo: memo,
            monthStock: mos,
            columns: columns
          });
        });
      } catch (e) {
        log.error('Bad custpage_selected JSON', e);
      }
    }

    if (!selectedRows.length) {
      Object.keys(params).forEach(function (key) {
        if (key.indexOf('row_select_') === 0) {
          var rowId = key.split('_')[2];
          if (!rows[rowId]) return;
          var qty = parseInt(params['qty_input_' + rowId], 10) || 0;
          var memo = params['memo_input_' + rowId] || '';
          var columns = parseCsvLine(rows[rowId]).map(function (val) {
            return cleanDisplayValue(val);
          });
          selectedRows.push({
            rowId: rowId,
            qty: qty,
            memo: memo,
            columns: columns
          });
        }
      });
    }

    selectedRows.forEach(function (entry) {
      var cols = entry.columns;
      var monStock = entry.monthStock;

      if (
        monStock == null ||
        monStock === '' ||
        monStock === 'null' ||
        monStock === 'Infinity' ||
        monStock === 'infinity' ||
        monStock === 'NaN' ||
        monStock === 'nan'
      ) {
        monStock = 0;
      }

      try {
        record.create({
          type: 'customrecord_mi_planned_po',
          isDynamic: true
        })
          .setValue({ fieldId: 'custrecord_mi_item', value: cols[3] })
          .setValue({ fieldId: 'custrecord_mi_order_qty', value: entry.qty })
          .setValue({ fieldId: 'custrecord_mi_purchase_memo', value: entry.memo === 0 ? '' : entry.memo })
          .setValue({ fieldId: 'custrecord_month_of_stocks', value: safeParseFloat(monStock) })
          .setValue({ fieldId: 'custrecord_mi_qty_of_ordered_not_ship', value: safeParseFloat(cols[24]) })
          .setValue({ fieldId: 'custrecord_mi_qty_available', value: safeParseFloat(cols[31]) })
          .setValue({ fieldId: 'custrecord_mi_qty_in_transit', value: safeParseFloat(cols[25]) })
          .setValue({ fieldId: 'custrecord_mi_min_month_qty', value: safeParseFloat(cols[16]) })
          .save();

        createdCount++;
      } catch (e) {
        log.error('Error creating custom record for row ' + entry.rowId, e);
      }
    });

    context.response.write(
      '<html>' +
      '<head>' +
      '<meta http-equiv="refresh" content="5;URL=' + escapeHtmlAttr(
        RETURN_URL_BASE +
        '&empid=' + encodeURIComponent(postedEmp) +
        '&ts=' + encodeURIComponent(postedTs) +
        '&sig=' + encodeURIComponent(postedSig)
      ) + '" />' +
      '<style>' +
      'body{font-family:Arial,sans-serif;text-align:center;padding-top:100px;}' +
      '.message-box{display:inline-block;background-color:#f3f9ff;padding:20px 30px;border:1px solid #a3d2f2;border-radius:8px;color:#005b99;font-size:16px;}' +
      '</style>' +
      '</head>' +
      '<body>' +
      '<div class="message-box"><strong>' + createdCount + '</strong> Planned Purchase Order(s) created.<br />You will be redirected to the main page in 5 seconds...</div>' +
      '</body>' +
      '</html>'
    );
  }

  function buildHeaderMeta(headers, typeParam) {
    var rawHeaders = headers.slice();
    var normalized = rawHeaders.map(function (h) {
      return normalizeHeader(h);
    });

    var adminCsvIndex = normalized.indexOf('admin portal');
    if (adminCsvIndex < 0) {
      for (var i = 0; i < normalized.length; i++) {
        if (
          normalized[i] === 'admin' ||
          normalized[i].indexOf('admin portal') === 0 ||
          normalized[i].indexOf('admin portal') >= 0
        ) {
          adminCsvIndex = i;
          break;
        }
      }
    }

    var truncateAfterAdmin = typeParam === '4';
    var removeJustAdmin = typeParam === '3';

    var headersForOutput;
    if (adminCsvIndex >= 0) {
      if (truncateAfterAdmin) {
        headersForOutput = rawHeaders.slice(0, adminCsvIndex);
      } else if (removeJustAdmin) {
        headersForOutput = rawHeaders.filter(function (_v, idx) {
          return idx !== adminCsvIndex;
        });
      } else {
        headersForOutput = rawHeaders.slice();
      }
    } else {
      headersForOutput = rawHeaders.slice();
    }

    headersForOutput.push('Filter');

    return {
      rawHeaders: rawHeaders,
      headersForOutput: headersForOutput,
      adminCsvIndex: adminCsvIndex,
      truncateAfterAdmin: truncateAfterAdmin,
      removeJustAdmin: removeJustAdmin
    };
  }

  function buildOutputRows(rows, headerMeta, balances) {
    var outputRows = [];
    outputRows.push(headerMeta.headersForOutput);

    var filterSets = {
      items: {},
      vendors: {},
      brands: {},
      brandCats: {},
      depts: {},
      pols: {},
      products: {}
    };

    var redRows = [];
    var blackRows = [];
    var alreadyExists = {};

    rows.slice(1).forEach(function (rowLine) {
      if (!rowLine || !String(rowLine).trim()) return;

      var columns = parseCsvLine(rowLine);
      if (!columns.length) return;

      if (headerMeta.truncateAfterAdmin && headerMeta.adminCsvIndex >= 0) {
        var adminCellVal = cleanDisplayValue(columns[headerMeta.adminCsvIndex]).toLowerCase();
        if (adminCellVal !== 'admin portal') return;
      }

      var calcCols = columns.slice();
      var monthAvg = safeParseFloat(calcCols[csvIndexIsExposed(11, headerMeta)]);
      var itemid = cleanDisplayValue(calcCols[csvIndexIsExposed(3, headerMeta)]);

      if (!itemid || alreadyExists[itemid]) return;
      alreadyExists[itemid] = true;

      addUnique(filterSets.items, cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.ITEM, headerMeta)]));
      addUnique(filterSets.vendors, cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.VENDOR, headerMeta)]));
      addUnique(filterSets.brands, cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.BRAND, headerMeta)]));
      addUnique(filterSets.brandCats, cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.BRAND_CATEGORY, headerMeta)]));
      addUnique(filterSets.depts, cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.DEPT, headerMeta)]));
      addUnique(filterSets.pols, cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.POL, headerMeta)]));
      addUnique(filterSets.products, cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.PRODUCT, headerMeta)]));

      var val43 = safeParseFloat(calcCols[csvIndexIsExposed(25, headerMeta)]);
      var val41 = safeParseFloat(calcCols[csvIndexIsExposed(20, headerMeta)]);
      var diff = Math.abs(val43 - val41);
      calcCols[csvIndexIsExposed(25, headerMeta)] = diff === 0 ? '' : String(diff);

      var good = 0;
      var bad = 0;
      var hold = 0;
      var inspect = 0;
      var label = 0;
      var total = 0;
      var avail = 0;

      if (balances[itemid]) {
        good = safeParseFloat(balances[itemid].good);
        bad = safeParseFloat(balances[itemid].bad);
        hold = safeParseFloat(balances[itemid].hold);
        inspect = safeParseFloat(balances[itemid].inspect);
        label = safeParseFloat(balances[itemid].label);
        total = safeParseFloat(balances[itemid].total);
        avail = safeParseFloat(balances[itemid].avail);
      }

      var col12Val = safeParseFloat(calcCols[csvIndexIsExposed(12, headerMeta)]);
      var availtoProm = good - col12Val;

      calcCols[calcCols.length] = 'Black';
      calcCols[csvIndexIsExposed(12, headerMeta)] = String(availtoProm);
      calcCols[csvIndexIsExposed(13, headerMeta)] = String(good);
      calcCols[csvIndexIsExposed(14, headerMeta)] = String(bad);
      calcCols[csvIndexIsExposed(15, headerMeta)] = String(inspect);
      calcCols[csvIndexIsExposed(16, headerMeta)] = String(label);
      calcCols[csvIndexIsExposed(17, headerMeta)] = String(hold);
      calcCols[csvIndexIsExposed(18, headerMeta)] = String(total);
      calcCols[csvIndexIsExposed(19, headerMeta)] = monthAvg ? ((safeParseFloat(total)) / monthAvg).toFixed(2) : '';
      calcCols[csvIndexIsExposed(24, headerMeta)] = monthAvg ? (((safeParseFloat(total)) + safeParseFloat(val41)) / monthAvg).toFixed(2) : '';
      calcCols[csvIndexIsExposed(23, headerMeta)] = (safeParseFloat(total) + safeParseFloat(val41)).toFixed(2);
      calcCols[csvIndexIsExposed(27, headerMeta)] = monthAvg ? ((safeParseFloat(total) + safeParseFloat(val43)) / monthAvg).toFixed(2) : '';
      calcCols[csvIndexIsExposed(26, headerMeta)] = (safeParseFloat(total) + safeParseFloat(val43)).toFixed(2);
      calcCols[csvIndexIsExposed(31, headerMeta)] = String(avail);

      var col9 = normalizeMovement(monthAvg);
      var qtytotal = diff + safeParseFloat(calcCols[csvIndexIsExposed(20, headerMeta)]) + safeParseFloat(avail) - safeParseFloat(col12Val);
      var stockingQty = Math.ceil(safeParseFloat(calcCols[csvIndexIsExposed(11, headerMeta)]) * 4.5);
      calcCols[csvIndexIsExposed(11, headerMeta)] = col9;

      calcCols[csvIndexIsExposed(63, headerMeta)] = String(calcCols[csvIndexIsExposed(11, headerMeta)]);
      calcCols[csvIndexIsExposed(64, headerMeta)] = String(safeParseFloat(calcCols[csvIndexIsExposed(11, headerMeta)]) * 4);

      var recommendedQty = 0;
      if (qtytotal < stockingQty) {
        recommendedQty = (stockingQty - qtytotal).toFixed(2);
      }

      var monthsStock = monthAvg ? ((diff + safeParseFloat(avail) + safeParseFloat(calcCols[csvIndexIsExposed(20, headerMeta)])) / monthAvg) : 0;

      if (calcCols[1] == 0) calcCols[1] = '';
      calcCols[1] = String(recommendedQty);

      var rowStyle = '';
      if (!isNaN(monthsStock) && monthsStock <= 4.5) {
        rowStyle = 'color:#c62828;';
        calcCols[calcCols.length - 1] = 'Red';
      }

      var statusVal = calcCols[calcCols.length - 1] || '';
      var baseCols = calcCols.slice(0, -1);
      var displayCols;

      if (headerMeta.adminCsvIndex >= 0 && headerMeta.truncateAfterAdmin) {
        displayCols = baseCols.slice(0, headerMeta.adminCsvIndex);
      } else if (headerMeta.adminCsvIndex >= 0 && headerMeta.removeJustAdmin) {
        displayCols = baseCols.filter(function (_v, idx) {
          return idx !== headerMeta.adminCsvIndex;
        });
      } else {
        displayCols = baseCols.slice();
      }
      displayCols.push(statusVal);

      var cleanedCols = displayCols.map(function (value) {
        return sanitizeCellForCsv(value);
      });
      outputRows.push(cleanedCols);

      var rowId = outputRows.length - 1;
      var rowHtml = buildTableRowHtml({
        rowId: rowId,
        rowStyle: rowStyle,
        recommendedQty: recommendedQty,
        monthsStock: monthsStock,
        displayCols: displayCols,
        item: cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.ITEM, headerMeta)]),
        vendor: cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.VENDOR, headerMeta)]),
        brand: cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.BRAND, headerMeta)]),
        brandCat: cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.BRAND_CATEGORY, headerMeta)]),
        dept: cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.DEPT, headerMeta)]),
        pol: cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.POL, headerMeta)]),
        product: cleanDisplayValue(columns[csvIndexIsExposed(HEADER_INDEX.PRODUCT, headerMeta)]),
        itemSpace: '',
        weight: ''
      });

      var bucket = rowStyle ? redRows : blackRows;
      bucket.push({
        itemId: safeParseInt(itemid),
        html: rowHtml
      });
    });

    redRows.sort(function (a, b) { return a.itemId - b.itemId; });
    blackRows.sort(function (a, b) { return a.itemId - b.itemId; });

    return {
      outputRows: outputRows,
      rowsHtml: redRows.map(function (r) { return r.html; }).join('') +
        blackRows.map(function (r) { return r.html; }).join(''),
      filterSets: {
        items: objectKeysSorted(filterSets.items),
        vendors: objectKeysSorted(filterSets.vendors),
        brands: objectKeysSorted(filterSets.brands),
        brandCats: objectKeysSorted(filterSets.brandCats),
        depts: objectKeysSorted(filterSets.depts),
        pols: objectKeysSorted(filterSets.pols),
        products: objectKeysSorted(filterSets.products)
      }
    };
  }

  function buildTableRowHtml(cfg) {
    var html = '';
    html += '<tr style="' + escapeHtmlAttr(cfg.rowStyle || '') + '"' +
      ' data-item="' + escapeHtmlAttr((cfg.item || '').toLowerCase()) + '"' +
      ' data-vendor="' + escapeHtmlAttr((cfg.vendor || '').toLowerCase()) + '"' +
      ' data-brand="' + escapeHtmlAttr((cfg.brand || '').toLowerCase()) + '"' +
      ' data-brandcat="' + escapeHtmlAttr((cfg.brandCat || '').toLowerCase()) + '"' +
      ' data-dept="' + escapeHtmlAttr((cfg.dept || '').toLowerCase()) + '"' +
      ' data-pol="' + escapeHtmlAttr((cfg.pol || '').toLowerCase()) + '"' +
      ' data-product="' + escapeHtmlAttr((cfg.product || '').toLowerCase()) + '"' +
      '>';

    html += '<td class="sticky-col c0"><input type="checkbox" name="row_select_' + cfg.rowId + '" /></td>';
    html += '<td class="sticky-col c1"><input type="number" class="qty-input" name="qty_input_' + cfg.rowId + '" min="0" value="' + escapeHtmlAttr(String(cfg.recommendedQty || 0)) + '" /></td>';
    html += '<td class="sticky-col c2 month-stock-cell">' + escapeHtml(isFinite(cfg.monthsStock) ? cfg.monthsStock.toFixed(2) : '') + '</td>';
    html += '<td class="sticky-col c3 item-space-cell">' + escapeHtml(cleanDisplayValue(cfg.itemSpace)) + '</td>';
    html += '<td class="sticky-col c4 weight-cell">' + escapeHtml(cleanDisplayValue(cfg.weight)) + '</td>';

    cfg.displayCols.forEach(function (value, idx) {
      var cleaned = cleanDisplayValue(value);
      if (idx === 0) {
        html += '<td class="sticky-col c5"><input type="text" class="memo-input" name="memo_input_' + cfg.rowId + '" value="' + escapeHtmlAttr(cleaned) + '" /></td>';
      } else if (idx < 5) {
        html += '<td class="sticky-col c' + (idx + 5) + '">' + escapeHtml(cleaned) + '</td>';
      } else {
        html += '<td>' + escapeHtml(cleaned) + '</td>';
      }
    });

    html += '</tr>';
    return html;
  }

  function buildHtml(cfg) {
    var headersHtml = '';
    headersHtml += '<th class="sticky-col c0">Select</th>';
    headersHtml += '<th class="sticky-col c1">Order Qty</th>';
    headersHtml += '<th class="sticky-col c2">Month of Stock</th>';
    headersHtml += '<th class="sticky-col c3">Item Space</th>';
    headersHtml += '<th class="sticky-col c4">Weight</th>';

    cfg.headersForOutput.forEach(function (value, idx) {
      var label = cleanDisplayValue(value) || ('Col ' + (idx + 1));
      var stickyClass = idx < 5 ? ' sticky-col c' + (idx + 5) : '';
      headersHtml += ''
        + '<th class="' + stickyClass + '" data-col="' + (idx + 5) + '">'
        +   '<div class="th-inner">'
        +     '<span class="th-text">' + escapeHtml(label) + '</span>'
        +     '<button type="button" class="th-filter-btn" data-col="' + (idx + 5) + '">▾</button>'
        +   '</div>'
        + '</th>';
    });


     var topFiltersHtml = '';
if (cfg.showTopFilters) {
  topFiltersHtml =
    '<div class="filter-shell">'
    + '<button id="topFilterBtn" class="btn" type="button">Filters</button>'
    + '<div id="topFilterPanel" class="filter-panel">'
      + '<div class="filter-grid">'
        + buildFilterBox('item', 'Item')
        + buildFilterBox('vendor', 'Vendor')
        + buildFilterBox('brand', 'Brand')
        + buildFilterBox('brandCat', 'Brand Category')
        + buildFilterBox('dept', 'Department')
        + buildFilterBox('pol', 'P.O.L')
        + buildFilterBox('product', 'Product Type')
      + '</div>'
      + '<div style="display:flex;justify-content:flex-end;gap:8px;margin-top:12px;">'
        + '<button type="button" class="btn" id="clearAllTopFilters">Clear All</button>'
        + '<button type="button" class="btn btn-primary" id="applyTopFilters">Apply</button>'
      + '</div>'
    + '</div>'
    + '</div>';
}

    return ''
      + '<style>'
      + 'body{font-family:Arial,sans-serif;}'
      + '.rr-wrap{width:100%;max-width:100%;box-sizing:border-box;}'
      + '.toolbar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:0 0 10px 0;flex-wrap:wrap;}'
      + '.toolbar-left,.toolbar-right{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}'
      + '.icon-btn,.btn{border:1px solid #cfd8e3;background:#fff;color:#1f2937;padding:7px 10px;border-radius:8px;cursor:pointer;font-size:12px;}'
      + '.icon-btn:hover,.btn:hover{background:#f8fafc;}'
      + '.btn-primary{background:#2563eb;border-color:#2563eb;color:#fff;}'
      + '.btn-primary:hover{background:#1d4ed8;}'
      + '.totals-bar{display:inline-flex;align-items:center;gap:8px;padding:7px 12px;border:1px solid #cfe3ff;border-radius:999px;background:#f4f8ff;}'
      + '.totals-label{font-size:11px;font-weight:700;color:#2b5a99;text-transform:uppercase;}'
      + '.totals-value{font-size:13px;font-weight:700;color:#0f3f78;font-variant-numeric:tabular-nums;}'
      + '.filter-shell{position:relative;}'
      + '.filter-panel{display:none;position:absolute;top:38px;left:0;z-index:9999;background:#fff;border:1px solid #d7dee8;border-radius:12px;box-shadow:0 12px 30px rgba(15,23,42,.12);width:980px;max-width:min(980px,95vw);padding:14px;}'
      + '.filter-panel.open{display:block;}'
      + '.filter-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;}'
      + '.filter-box{border:1px solid #e5e7eb;border-radius:10px;padding:10px;background:#fbfdff;}'
      + '.filter-box label{display:block;font-size:12px;font-weight:700;margin:0 0 6px 0;color:#334155;}'
      + '.filter-search{width:100%;box-sizing:border-box;padding:7px 8px;border:1px solid #cbd5e1;border-radius:8px;margin-bottom:8px;font-size:12px;}'
      + '.filter-list{border:1px solid #e5e7eb;border-radius:8px;height:170px;overflow:auto;background:#fff;padding:6px;}'
      + '.filter-row{display:flex;align-items:center;gap:8px;padding:4px 2px;font-size:12px;}'
      + '.filter-actions{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px;}'
      + '.table-wrap{width:100%;max-width:calc(100vw - 24px);border:1px solid #d7dee8;border-radius:12px;overflow:hidden;background:#fff;}'
      + '.top-scroll{height:14px;overflow-x:auto;overflow-y:hidden;border-bottom:1px solid #e5e7eb;background:#f8fafc;}'
      + '.top-scroll-inner{height:1px;}'
      + '.table-container{max-height:78vh;overflow:auto;position:relative;}'
      + '#excelTable{border-collapse:separate;border-spacing:0;width:max-content;min-width:100%;table-layout:fixed;}'
      + '#excelTable th,#excelTable td{border-right:1px solid #d7dee8;border-bottom:1px solid #d7dee8;padding:7px 10px;background:#fff;white-space:nowrap;font-size:12px;box-sizing:border-box;}'
      + '#excelTable th{position:sticky;top:0;background:#f6f8fb;z-index:40;font-weight:700;color:#1f2937;}'
      + '#excelTable tr:nth-child(even) td{background:#fcfdff;}'
      + '#excelTable tr:hover td{background:#f8fbff;}'
      + '#excelTable th:first-child,#excelTable td:first-child{border-left:1px solid #d7dee8;}'
      + '#excelTable thead tr:first-child th{border-top:1px solid #d7dee8;}'
      + '.th-inner{display:flex;align-items:center;justify-content:space-between;gap:6px;}'
      + '.th-text{display:inline-block;overflow:hidden;text-overflow:ellipsis;}'
      + '.th-filter-btn{border:1px solid #d1d5db;background:#fff;border-radius:6px;padding:2px 6px;cursor:pointer;font-size:11px;line-height:1.2;}'
      + '.th-filter-btn:hover{background:#f3f4f6;}'
      + '.header-panel{position:fixed;z-index:10000;width:280px;max-height:360px;overflow:hidden;background:#fff;border:1px solid #cbd5e1;border-radius:12px;box-shadow:0 12px 30px rgba(15,23,42,.18);padding:10px;}'
      + '.header-panel .title{font-size:12px;font-weight:700;margin-bottom:8px;}'
      + '.header-panel .list{border:1px solid #e5e7eb;border-radius:8px;max-height:210px;overflow:auto;padding:6px;background:#fff;}'
      + '.header-panel .row{display:flex;align-items:center;gap:8px;padding:4px 2px;font-size:12px;}'
      + '.header-panel .actions{display:flex;justify-content:space-between;gap:6px;margin-top:8px;flex-wrap:wrap;}'
      + '.sticky-col{position:sticky;background:#fff;z-index:30;box-shadow:inset -1px 0 0 #d7dee8;}'
      + '#excelTable thead .sticky-col{z-index:60;background:#f6f8fb;}'
      + '.c0{left:0;min-width:58px;max-width:58px;width:58px;}'
      + '.c1{left:58px;min-width:105px;max-width:105px;width:105px;}'
      + '.c2{left:163px;min-width:118px;max-width:118px;width:118px;}'
      + '.c3{left:281px;min-width:110px;max-width:110px;width:110px;}'
      + '.c4{left:391px;min-width:100px;max-width:100px;width:100px;}'
      + '.c5{left:491px;min-width:220px;max-width:220px;width:220px;}'
      + '.c6{left:711px;min-width:120px;max-width:120px;width:120px;}'
      + '.c7{left:831px;min-width:170px;max-width:170px;width:170px;}'
      + '.c8{left:1001px;min-width:110px;max-width:110px;width:110px;}'
      + '.c9{left:1111px;min-width:420px;max-width:420px;width:420px;}'
      + '.qty-input,.memo-input{width:100%;box-sizing:border-box;padding:5px 6px;border:1px solid #cbd5e1;border-radius:6px;font-size:12px;background:#fff;}'
      + '.qty-input{min-width:80px;}'
      + '.hidden-row{display:none !important;}'
      + '.muted{color:#64748b;font-size:11px;}'
      + '.panel-backdrop{display:none;}'
      + '@media (max-width:980px){.filter-grid{grid-template-columns:1fr;}.filter-panel{width:min(95vw,980px);}}'
      + '</style>'

      + '<div class="rr-wrap">'
      +   '<div class="toolbar">'
      +     '<div class="toolbar-left">'
      +       '<button id="downloadCsvBtn" class="icon-btn" type="button" title="Download CSV">Download CSV</button>'
      +   topFiltersHtml
      +     '</div>'
      +     '<div class="toolbar-right">'
      +       '<div class="totals-bar"><span class="totals-label">Selected Rows</span><span id="selectedCount" class="totals-value">0</span></div>'
      +       '<div class="totals-bar"><span class="totals-label">Visible Rows</span><span id="visibleCount" class="totals-value">0</span></div>'
      +       '<div class="totals-bar"><span class="totals-label">Total Cubic Space</span><span id="totalCubicValue" class="totals-value">0</span></div>'
      +       '<div class="totals-bar"><span class="totals-label">Total Weight</span><span id="totalWeightValue" class="totals-value">0</span></div>'
      +     '</div>'
      +   '</div>'

      +   '<div class="table-wrap">'
      +     '<div id="topScroll" class="top-scroll"><div id="topScrollInner" class="top-scroll-inner"></div></div>'
      +     '<div id="tableContainer" class="table-container">'
      +       '<table id="excelTable">'
      +         '<thead><tr>' + headersHtml + '</tr></thead>'
      +         '<tbody>' + cfg.rowsHtml + '</tbody>'
      +       '</table>'
      +     '</div>'
      +   '</div>'
      + '</div>'

      + '<div id="filterPortal"></div>'

      + '<script>'
      + 'window.DOWNLOAD_URL=' + JSON.stringify(cfg.downloadUrl || '') + ';'
      + 'window.__FILTER_DATA__=' + JSON.stringify(cfg.filterSets) + ';'
      + 'window.__ADMIN_IDX__=' + JSON.stringify(cfg.adminCsvIndex) + ';'
      + 'window.__TRUNC_AFTER__=' + JSON.stringify(cfg.truncateAfterAdmin && cfg.adminCsvIndex >= 0) + ';'
      + 'window.__ADMIN_REMOVED__=' + JSON.stringify(cfg.removeJustAdmin && cfg.adminCsvIndex >= 0) + ';'
      + '</script>'

      + '<script>'
      + '(function(){'
      + 'function qs(s,r){return (r||document).querySelector(s);}'
      + 'function qsa(s,r){return Array.prototype.slice.call((r||document).querySelectorAll(s));}'
      + 'function text(v){return String(v==null?"":v);}'
      + 'function lower(v){return text(v).toLowerCase();}'
      + 'function byId(id){return document.getElementById(id);}'
      + 'function uniqSorted(arr){return arr.slice().sort(function(a,b){return String(a).localeCompare(String(b),undefined,{numeric:true,sensitivity:"base"});});}'
      + 'function toNumber(v){var n=parseFloat(v);return isNaN(n)?0:n;}'
      + 'function escapeHtml(v){return text(v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/\'/g,"&#39;");}'

      + 'var table=byId("excelTable");'
      + 'var container=byId("tableContainer");'
      + 'var topScroll=byId("topScroll");'
      + 'var topInner=byId("topScrollInner");'
      + 'var suiteletForm=document.querySelector("form");'
      + 'var hiddenSelected=byId("custpage_selected");'
      + 'var topFilterBtn=byId("topFilterBtn");'
      + 'var topFilterPanel=byId("topFilterPanel");'
      + 'var filterData=window.__FILTER_DATA__||{};'
      + 'var selectedCountEl=byId("selectedCount");'
      + 'var visibleCountEl=byId("visibleCount");'
      + 'var totalCubicValue=byId("totalCubicValue");'
      + 'var totalWeightValue=byId("totalWeightValue");'
      + 'var activeHeaderFilters={};'
      + 'var activeTopFilters={ item:new Set(), vendor:new Set(), brand:new Set(), brandCat:new Set(), dept:new Set(), pol:new Set(), product:new Set() };'

      + 'function syncTopScroll(){ if(!table||!topInner)return; topInner.style.width=Math.max(table.scrollWidth, container.clientWidth+2)+"px"; }'
      + 'if(topScroll&&container){ topScroll.addEventListener("scroll",function(){ container.scrollLeft=topScroll.scrollLeft; }); container.addEventListener("scroll",function(){ topScroll.scrollLeft=container.scrollLeft; }); }'
      + 'window.addEventListener("resize", function(){ setTimeout(syncTopScroll, 50); });'
      + 'setTimeout(syncTopScroll, 50); setTimeout(syncTopScroll, 250);'

      + 'function getBodyRows(){ return qsa("#excelTable tbody tr"); }'
      + 'function isRowVisible(tr){ return tr.style.display !== "none"; }'
      + 'function updateSummary(){'
      + '  var rows=getBodyRows();'
      + '  var visible=0, selected=0, cubic=0, weight=0;'
      + '  rows.forEach(function(tr){'
      + '    if(isRowVisible(tr)) visible++;'
      + '    var cb=qs(\'input[type="checkbox"][name^="row_select_"]\', tr);'
      + '    if(cb && cb.checked){'
      + '      selected++;'
      + '      cubic += toNumber((qs(".item-space-cell", tr)||{}).textContent);'
      + '      weight += toNumber((qs(".weight-cell", tr)||{}).textContent);'
      + '    }'
      + '  });'
      + '  if(selectedCountEl) selectedCountEl.textContent=String(selected);'
      + '  if(visibleCountEl) visibleCountEl.textContent=String(visible);'
      + '  if(totalCubicValue) totalCubicValue.textContent=(Math.round(cubic*100)/100).toFixed(2).replace(/\\.00$/,"");'
      + '  if(totalWeightValue) totalWeightValue.textContent=(Math.round(weight*100)/100).toFixed(2).replace(/\\.00$/,"");'
      + '}'

      + 'function createFilterOptions(boxId, values){'
      + '  var list=byId(boxId+"_list");'
      + '  var search=byId(boxId+"_search");'
      + '  if(!list) return;'
      + '  values=uniqSorted(values||[]);'
      + '  function render(term){'
      + '    var t=lower(term);'
      + '    list.innerHTML="";'
      + '    values.forEach(function(v){'
      + '      if(t && lower(v).indexOf(t)===-1) return;'
      + '      var row=document.createElement("div");'
      + '      row.className="filter-row";'
      + '      row.innerHTML=\'<input type="checkbox" value="\'+escapeHtml(v)+\'"><span>\'+escapeHtml(v||"(blank)")+\'</span>\';'
      + '      list.appendChild(row);'
      + '    });'
      + '    restoreTopSelections(boxId);'
      + '  }'
      + '  if(search){ search.addEventListener("input", function(){ render(this.value||""); }); }'
      + '  render("");'
      + '}'

      + 'function restoreTopSelections(boxId){'
      + '  var map={ item:"item", vendor:"vendor", brand:"brand", brandCat:"brandCat", dept:"dept", pol:"pol", product:"product" };'
      + '  var key=map[boxId]; if(!key) return;'
      + '  var selected=activeTopFilters[key];'
      + '  qsa("input[type=checkbox]", byId(boxId+"_list")).forEach(function(cb){'
      + '    cb.checked = selected.has(lower(cb.value));'
      + '  });'
      + '}'

      + 'function syncSelectionsFromBox(boxId){'
      + '  var map={ item:"item", vendor:"vendor", brand:"brand", brandCat:"brandCat", dept:"dept", pol:"pol", product:"product" };'
      + '  var key=map[boxId]; if(!key) return;'
      + '  var selected=activeTopFilters[key];'
      + '  qsa("input[type=checkbox]", byId(boxId+"_list")).forEach(function(cb){'
      + '    if(cb.checked) selected.add(lower(cb.value)); else selected.delete(lower(cb.value));'
      + '  });'
      + '}'

      + '["item","vendor","brand","brandCat","dept","pol","product"].forEach(function(key){'
      + '  createFilterOptions(key, filterData[key+"s"] || []);'
      + '  var list=byId(key+"_list");'
      + '  if(list){ list.addEventListener("change", function(){ syncSelectionsFromBox(key); }); }'
      + '  var allBtn=byId(key+"_selectall");'
      + '  var noneBtn=byId(key+"_deselectall");'
      + '  if(allBtn){ allBtn.addEventListener("click", function(){ qsa("input[type=checkbox]", list).forEach(function(cb){ cb.checked=true; }); syncSelectionsFromBox(key); }); }'
      + '  if(noneBtn){ noneBtn.addEventListener("click", function(){ qsa("input[type=checkbox]", list).forEach(function(cb){ cb.checked=false; }); syncSelectionsFromBox(key); }); }'
      + '});'

      + 'if(topFilterBtn && topFilterPanel){'
      + '  topFilterBtn.addEventListener("click", function(e){ e.stopPropagation(); topFilterPanel.classList.toggle("open"); });'
      + '  document.addEventListener("click", function(e){ if(!topFilterPanel.contains(e.target) && e.target!==topFilterBtn){ topFilterPanel.classList.remove("open"); } });'
      + '}'

      + 'function rowPassTopFilters(tr){'
      + '  function ok(key, attr){ var set=activeTopFilters[key]; if(!set || !set.size) return true; return set.has(lower(tr.getAttribute(attr)||"")); }'
      + '  return ok("item","data-item") && ok("vendor","data-vendor") && ok("brand","data-brand") && ok("brandCat","data-brandcat") && ok("dept","data-dept") && ok("pol","data-pol") && ok("product","data-product");'
      + '}'

      + 'function getCellText(tr, idx){ var td=tr.cells[idx]; return td ? lower(td.textContent||"") : ""; }'
      + 'function rowPassHeaderFilters(tr){'
      + '  for(var k in activeHeaderFilters){'
      + '    if(!activeHeaderFilters.hasOwnProperty(k)) continue;'
      + '    var set=activeHeaderFilters[k];'
      + '    if(set && set.size){ if(!set.has(getCellText(tr, parseInt(k,10)))) return false; }'
      + '  }'
      + '  return true;'
      + '}'

      + 'function applyAllFilters(){'
      + '  getBodyRows().forEach(function(tr){'
      + '    tr.style.display = (rowPassTopFilters(tr) && rowPassHeaderFilters(tr)) ? "" : "none";'
      + '  });'
      + '  updateSummary();'
      + '}'

      + 'var applyTopFiltersBtn=byId("applyTopFilters");'
      + 'if(applyTopFiltersBtn){ applyTopFiltersBtn.addEventListener("click", function(){ applyAllFilters(); if(topFilterPanel) topFilterPanel.classList.remove("open"); }); }'
      + 'var clearAllTopFiltersBtn=byId("clearAllTopFilters");'
      + 'if(clearAllTopFiltersBtn){ clearAllTopFiltersBtn.addEventListener("click", function(){'
      + '  Object.keys(activeTopFilters).forEach(function(k){ activeTopFilters[k]=new Set(); });'
      + '  ["item","vendor","brand","brandCat","dept","pol","product"].forEach(function(key){'
      + '    qsa("input[type=checkbox]", byId(key+"_list")).forEach(function(cb){ cb.checked=false; });'
      + '    var s=byId(key+"_search"); if(s) s.value="";'
      + '    createFilterOptions(key, filterData[key+"s"] || []);'
      + '  });'
      + '  applyAllFilters();'
      + '}); }'

      + 'function closeHeaderPanel(){ var p=qs(".header-panel"); if(p){ p.remove(); } }'
      + 'function getColumnUniverse(colIdx){'
      + '  var seen={};'
      + '  getBodyRows().forEach(function(tr){'
      + '    var val=(tr.cells[colIdx]?tr.cells[colIdx].textContent:"")||"";'
      + '    val=val===" - None - " ? "" : val;'
      + '    seen[lower(val.trim())]=val.trim();'
      + '  });'
      + '  return Object.keys(seen).map(function(k){ return { key:k, label: seen[k] || "(blank)" }; }).sort(function(a,b){ return a.label.localeCompare(b.label,undefined,{numeric:true,sensitivity:"base"}); });'
      + '}'

      + 'function openHeaderFilter(btn, colIdx){'
      + '  closeHeaderPanel();'
      + '  var panel=document.createElement("div");'
      + '  panel.className="header-panel";'
      + '  panel.innerHTML='
      + '    \'<div class="title">Filter</div>\' +'
      + '    \'<input type="search" class="filter-search" placeholder="Search values...">\' +'
      + '    \'<div class="list"></div>\' +'
      + '    \'<div class="actions">\' +'
      + '      \'<button type="button" class="btn" data-act="clear">Clear</button>\' +'
      + '      \'<button type="button" class="btn" data-act="selectall">Select all</button>\' +'
      + '      \'<button type="button" class="btn" data-act="deselectall">Deselect all</button>\' +'
      + '      \'<button type="button" class="btn btn-primary" data-act="apply">Apply</button>\' +'
      + '    \'</div>\';'
      + '  var list=qs(".list", panel);'
      + '  var search=qs(".filter-search", panel);'
      + '  var universe=getColumnUniverse(colIdx);'
      + '  var current=activeHeaderFilters[colIdx] ? new Set(Array.from(activeHeaderFilters[colIdx])) : new Set(universe.map(function(v){ return v.key; }));'
      + '  function render(term){'
      + '    var t=lower(term); list.innerHTML="";'
      + '    universe.forEach(function(row){'
      + '      if(t && lower(row.label).indexOf(t)===-1) return;'
      + '      var item=document.createElement("div");'
      + '      item.className="row";'
      + '      item.setAttribute("data-key", row.key);'
      + '      item.innerHTML=\'<input type="checkbox" \'+(current.has(row.key)?\'checked\':\'\')+\'> <span>\'+escapeHtml(row.label||"(blank)")+\'</span>\';'
      + '      list.appendChild(item);'
      + '    });'
      + '  }'
      + '  render("");'
      + '  search.addEventListener("input", function(){ render(this.value||""); });'
      + '  list.addEventListener("change", function(e){'
      + '    var row=e.target.closest(".row"); if(!row) return;'
      + '    var key=row.getAttribute("data-key")||"";'
      + '    if(e.target.checked) current.add(key); else current.delete(key);'
      + '  });'
      + '  qs(".actions", panel).addEventListener("click", function(e){'
      + '    var act=e.target.getAttribute("data-act"); if(!act) return;'
      + '    if(act==="clear"){ delete activeHeaderFilters[colIdx]; closeHeaderPanel(); applyAllFilters(); return; }'
      + '    if(act==="selectall"){ qsa("input[type=checkbox]", list).forEach(function(cb){ cb.checked=true; current.add(cb.closest(".row").getAttribute("data-key")||""); }); return; }'
      + '    if(act==="deselectall"){ qsa("input[type=checkbox]", list).forEach(function(cb){ cb.checked=false; current.delete(cb.closest(".row").getAttribute("data-key")||""); }); return; }'
      + '    if(act==="apply"){'
      + '      if(current.size===universe.length){ delete activeHeaderFilters[colIdx]; } else { activeHeaderFilters[colIdx]=new Set(Array.from(current)); }'
      + '      closeHeaderPanel();'
      + '      applyAllFilters();'
      + '    }'
      + '  });'
      + '  document.body.appendChild(panel);'
      + '  var rect=btn.getBoundingClientRect();'
      + '  var top=rect.bottom+6;'
      + '  var left=Math.max(8, Math.min(window.innerWidth - 300, rect.right - 280));'
      + '  panel.style.top=top+"px"; panel.style.left=left+"px";'
      + '  setTimeout(function(){'
      + '    document.addEventListener("click", function handler(ev){'
      + '      if(panel.contains(ev.target) || btn.contains(ev.target)) return;'
      + '      panel.remove();'
      + '      document.removeEventListener("click", handler);'
      + '    });'
      + '  },0);'
      + '}'

      + 'if(table && table.tHead){'
      + '  table.tHead.addEventListener("click", function(e){'
      + '    var btn=e.target.closest(".th-filter-btn");'
      + '    if(!btn) return;'
      + '    e.preventDefault(); e.stopPropagation();'
      + '    openHeaderFilter(btn, parseInt(btn.getAttribute("data-col"),10));'
      + '  });'
      + '}'

      + 'var downloadBtn=byId("downloadCsvBtn");'
      + 'if(downloadBtn){ downloadBtn.addEventListener("click", function(e){ e.preventDefault(); if(window.DOWNLOAD_URL){ window.open(window.DOWNLOAD_URL, "_blank"); } else { alert("Download not available yet."); } }); }'

      + 'function collectSelectedRows(){'
      + '  if(!hiddenSelected) return;'
      + '  var out=[];'
      + '  getBodyRows().forEach(function(tr){'
      + '    var cb=qs(\'input[type="checkbox"][name^="row_select_"]\', tr);'
      + '    if(cb && cb.checked){'
      + '      var rowId=(cb.name||"").split("_").pop();'
      + '      var qtyInput=qs(\'input[type="number"][name^="qty_input_"]\', tr);'
      + '      var memoInput=qs(\'input[type="text"][name^="memo_input_"]\', tr);'
      + '      var mosCell=qs(".month-stock-cell", tr);'
      + '      out.push({'
      + '        rowId: rowId,'
      + '        qty: qtyInput ? (qtyInput.value||"0") : "0",'
      + '        memo: memoInput ? (memoInput.value||"") : "",'
      + '        mos: mosCell ? (mosCell.textContent||"").trim() : ""'
      + '      });'
      + '    }'
      + '  });'
      + '  hiddenSelected.value=JSON.stringify(out);'
      + '}'

      + 'function nativeSubmit(){'
      + '  if(!suiteletForm) return;'
      + '  collectSelectedRows();'
      + '  closeHeaderPanel();'
      + '  try{ HTMLFormElement.prototype.submit.call(suiteletForm); }'
      + '  catch(e){ if(typeof suiteletForm.submit==="function") suiteletForm.submit(); }'
      + '}'

      + 'var nsSubmitBtn=document.querySelector(\'input[type="submit"],button[type="submit"]\');'
      + 'if(nsSubmitBtn){ nsSubmitBtn.addEventListener("click", function(){ setTimeout(nativeSubmit, 0); }, true); }'
      + 'if(suiteletForm){ suiteletForm.addEventListener("submit", function(){ setTimeout(collectSelectedRows, 0); }, true); }'

      + 'document.addEventListener("change", function(e){'
      + '  if(e.target && (e.target.matches(\'input[type="checkbox"][name^="row_select_"]\') || e.target.matches(".qty-input"))){ updateSummary(); }'
      + '});'

      + 'applyAllFilters();'
      + 'updateSummary();'
      + '})();'
      + '</script>';
  }

  function buildFilterBox(id, label) {
    return ''
      + '<div class="filter-box">'
      +   '<label for="' + id + '_search">' + escapeHtml(label) + '</label>'
      +   '<input type="search" id="' + id + '_search" class="filter-search" placeholder="Search ' + escapeHtmlAttr(label) + '">'
      +   '<div id="' + id + '_list" class="filter-list"></div>'
      +   '<div class="filter-actions">'
      +     '<button type="button" class="btn" id="' + id + '_selectall">Select All</button>'
      +     '<button type="button" class="btn" id="' + id + '_deselectall">Deselect All</button>'
      +   '</div>'
      + '</div>';
  }

  function getLatestFiles() {
    var out = {
      mainId: null,
      secondaryId: null,
      thirdId: null
    };

    var folderSearchObj = search.create({
      type: 'folder',
      filters: [
        ['internalid', 'anyof', String(SOURCE_FOLDERS.MAIN), String(SOURCE_FOLDERS.SECONDARY), String(SOURCE_FOLDERS.THIRD)],
        'AND',
        ['file.documentsize', 'greaterthan', '5']
      ],
      columns: [
        search.createColumn({ name: 'internalid', summary: 'GROUP' }),
        search.createColumn({ name: 'internalid', join: 'file', summary: 'MAX' })
      ]
    });

    folderSearchObj.run().each(function (result) {
      var folderId = result.getValue({ name: 'internalid', summary: 'GROUP' });
      var latestFileId = result.getValue({ name: 'internalid', join: 'file', summary: 'MAX' });

      if (String(folderId) === String(SOURCE_FOLDERS.MAIN)) out.mainId = latestFileId;
      if (String(folderId) === String(SOURCE_FOLDERS.SECONDARY)) out.secondaryId = latestFileId;
      if (String(folderId) === String(SOURCE_FOLDERS.THIRD)) out.thirdId = latestFileId;

      return true;
    });

    return out;
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
        search.createColumn({
          name: 'formulanumeric',
          summary: 'SUM',
          formula: "case when {status} = 'Good' then {onhand} else 0 end"
        }),
        search.createColumn({
          name: 'formulanumeric1',
          summary: 'SUM',
          formula: "case when {status} = 'Deviation' then {onhand} else 0 end"
        }),
        search.createColumn({
          name: 'formulanumeric2',
          summary: 'SUM',
          formula: "case when {status} = 'Hold' then {onhand} else 0 end"
        }),
        search.createColumn({
          name: 'formulanumeric3',
          summary: 'SUM',
          formula: "case when {status} = 'Inspection' then {onhand} else 0 end"
        }),
        search.createColumn({
          name: 'formulanumeric4',
          summary: 'SUM',
          formula: "case when {status} = 'Label' then {onhand} else 0 end"
        }),
        search.createColumn({ name: 'available', summary: 'SUM', label: 'Available' })
      ]
    });

    inventorybalanceSearchObj.run().each(function (result) {
      var itemId = result.getValue({ name: 'item', summary: 'GROUP' });
      var goodQty = safeParseFloat(result.getValue({ name: 'formulanumeric', summary: 'SUM' }));
      var badQty = safeParseFloat(result.getValue({ name: 'formulanumeric1', summary: 'SUM' }));
      var holdQty = safeParseFloat(result.getValue({ name: 'formulanumeric2', summary: 'SUM' }));
      var inspectQty = safeParseFloat(result.getValue({ name: 'formulanumeric3', summary: 'SUM' }));
      var labelQty = safeParseFloat(result.getValue({ name: 'formulanumeric4', summary: 'SUM' }));
      var total = safeParseFloat((goodQty + badQty).toFixed(2));
      var avail = safeParseFloat(result.getValue({ name: 'available', summary: 'SUM' }));

      resultMap[itemId] = {
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

  function buildCsvContent(outputRows) {
    var lines = [];
    outputRows.forEach(function (row) {
      lines.push(row.map(function (cell) {
        return sanitizeCellForCsv(cell);
      }).join(','));
    });
    return '\uFEFF' + lines.join('\n');
  }

  function sanitizeCellForCsv(value) {
    var v = sanitizeText(value);
    v = cleanDisplayValue(v);

    if (v === '.00') v = '0.00';

    var needsQuotes = /[",\r\n]/.test(v);
    if (needsQuotes) {
      v = '"' + v.replace(/"/g, '""') + '"';
    }
    return v;
  }

  function cleanDisplayValue(value) {
    var v = value == null ? '' : String(value);
    v = v.replace(/^"+|"+$/g, '');
    v = v.replace(/\uFEFF/g, '');
    v = v.trim();
    if (v === '- None -' || v === 'NaN' || v === 'Infinity' || v === 'undefined' || v === 'null') v = '';
    if (v === '.00') v = '0.00';
    return v;
  }

  function sanitizeText(value) {
    var s = value == null ? '' : String(value);
    s = s.replace(/\uFEFF/g, '');
    s = s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '');
    s = s.replace(/\r\n/g, '\n');
    s = s.replace(/\r/g, '\n');
    s = s.replace(/[“”]/g, '"');
    s = s.replace(/[‘’]/g, "'");
    s = s.replace(/[–—]/g, '-');
    s = s.replace(/\u00A0/g, ' ');
    return s;
  }

  function splitCsvLines(content) {
    var txt = sanitizeText(content || '');
    if (!txt) return [];
    var parts = txt.split('\n');
    var lines = [];
    for (var i = 0; i < parts.length; i++) {
      if (parts[i] != null) lines.push(parts[i]);
    }
    return lines;
  }

  function parseCsvLine(line) {
    var out = [];
    var cur = '';
    var inQuotes = false;
    var i;
    var s = line == null ? '' : String(line);

    for (i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      var next = s.charAt(i + 1);

      if (ch === '"') {
        if (inQuotes && next === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === ',' && !inQuotes) {
        out.push(cur);
        cur = '';
      } else {
        cur += ch;
      }
    }
    out.push(cur);
    return out;
  }

  function normalizeHeader(h) {
    return String(h || '')
      .replace(/^[\uFEFF\s"]+|[\s"]+$/g, '')
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  function csvIndexIsExposed(idx, headerMeta) {
    if (headerMeta.adminCsvIndex >= 0 && idx >= headerMeta.adminCsvIndex) {
      return idx + 1;
    }
    return idx;
  }

  function normalizeMovement(val) {
    if (val === null || val === undefined) return 'No Movement';
    var s = String(val).replace(/,/g, '').trim();
    if (s === '') return 'No Movement';

    var n = parseFloat(s);
    if (!isFinite(n)) return 'No Movement';
    if (n <= 0) return 'No Movement';
    return n.toFixed(2);
  }

  function signCron(ts) {
    var h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
    h.update({ input: 'CRON|' + ts + '|rR9Z7KpXw2N6C8mE4HqFJYvT5bS0aUeD1LQG3oM' });
    return h.digest({ outputEncoding: crypto.Encoding.HEX });
  }

  function verifyCron(ts, sig) {
    if (!ts || !sig) return false;
    if (Math.abs(Date.now() - parseInt(ts, 10)) > TOKEN_TTL_MS) return false;
    try {
      return signCron(ts) === sig;
    } catch (e) {
      log.error('verifyCron token', e);
      return false;
    }
  }

  function sign(empid, ts) {
    var h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
    h.update({ input: empid + '|' + ts + '|' + SECRET });
    return h.digest({ outputEncoding: crypto.Encoding.HEX });
  }

  function verify(empid, ts, sig) {
    if (!empid || !ts || !sig) return false;
    if (Math.abs(Date.now() - parseInt(ts, 10)) > TOKEN_TTL_MS) return false;
    try {
      return sign(empid, ts) === sig;
    } catch (e) {
      log.error('verify token', e);
      return false;
    }
  }

  function addUnique(obj, val) {
    var v = cleanDisplayValue(val);
    if (!v) return;
    obj[v] = true;
  }

  function objectKeysSorted(obj) {
    return Object.keys(obj).sort(function (a, b) {
      return String(a).localeCompare(String(b), undefined, {
        numeric: true,
        sensitivity: 'base'
      });
    });
  }

  function safeParseFloat(val) {
    var num = parseFloat(val);
    return isNaN(num) ? 0 : num;
  }

  function safeParseInt(val) {
    var num = parseInt(val, 10);
    return isNaN(num) ? 0 : num;
  }

  function sanitizeFileName(name) {
    var s = sanitizeText(name || 'download.csv');
    s = s.replace(/[\\/:*?"<>|]+/g, '_');
    return s || 'download.csv';
  }

  function escapeHtml(str) {
    return String(str == null ? '' : str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function escapeHtmlAttr(str) {
    return escapeHtml(str).replace(/\n/g, '&#10;');
  }

  function writeLoginRequired(response) {
    response.write(
      '<html><head>' +
      '<script>setTimeout(function(){ window.location.href = ' + JSON.stringify(PORTAL_URL) + '; }, 1200);</script>' +
      '<style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>' +
      '</head><body><div class="message">Login Required</div></body></html>'
    );
  }

  function writeSessionExpired(response) {
    response.write(
      '<html><head>' +
      '<script>setTimeout(function(){ window.location.href = ' + JSON.stringify(PORTAL_URL) + '; }, 1200);</script>' +
      '<style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>' +
      '</head><body><div class="message">Session expired. Please log in again.</div></body></html>'
    );
  }

  return {
    onRequest: onRequest
  };
});