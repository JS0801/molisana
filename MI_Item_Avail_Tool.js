/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/ui/serverWidget', 'N/file', 'N/log', 'N/search', 'N/runtime', 'N/crypto'],
function (ui, file, log, search, runtime, crypto) {
  
  // ====== CONFIG ======
  var PORTAL_URL = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2110&deploy=1&compid=4975346&ns-at=AAEJ7tMQamzukv1WMqTK6i2c27bRetbrd2MDLjhDgPPFOawMxCo';
  
  var FOLDER_ITEM_LIST     = 423668;
  var FOLDER_INV_ITEM_LIST = 423667;
  var FOLDER_ASSEMBLY      = 423666;
  var FOLDER_DOWNLOAD_UI   = 423669;
  var FOLDER_DOWNLOAD_CRON = 279208;
  var lastBilledFile       = 447164;
  
  const TOKEN_TTL_MS = 30 * 60 * 1000;
  
  function sign(empid, ts) {
    var SECRET = runtime.getCurrentScript().getParameter({ name: 'custscript_portal_secret' }) || 'change-me';
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

  function signCron(ts, secret) {
    var h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
    h.update({ input: 'CRON|' + ts + '|' + secret });
    return h.digest({ outputEncoding: crypto.Encoding.HEX });
  }

  function verifyCron(ts, sig) {
    if (!ts || !sig) return false;
    if (Math.abs(Date.now() - parseInt(ts, 10)) > TOKEN_TTL_MS) return false;
    try {
      var secret = 'rR9Z7KpXw2N6C8mE4HqFJYvT5bS0aUeD1LQG3oM';
      return signCron(ts, secret) === sig;
    } catch (e) {
      log.error('verifyCron error', e);
      return false;
    }
  }

  function splitCsvRow(line) {
    if (!line && line !== '') return [];
    return line.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/);
  }

  function cleanHeader(h) {
    return String(h || '')
      .replace(/^[\uFEFF\s"]+|[\s"]+$/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function cleanCsvValue(v) {
    var s = (v == null ? '' : String(v).trim());
    if (s.length >= 2 && s[0] === '"' && s[s.length - 1] === '"') {
      var inner = s.slice(1, -1);
      if (!/[",\r\n]/.test(inner)) {
        s = inner;
      }
    }
    if (/[",\r\n]/.test(s)) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    if (s === '- None -' || s === 'NaN' || s === 'Infinity') s = '';
    return s;
  }

  function loadCsvRows(fileId) {
    if (!fileId) return [];
    var f = file.load({ id: fileId });
    var content = f.getContents() || '';
    var rows = content.split('\n');
    rows = rows.map(function (r) { return r.replace(/\r$/, ''); });
    return rows;
  }

  function toNumber(val) {
    var n = parseFloat(String(val == null ? '' : val).replace(/,/g, '').replace(/"/g, '').trim());
    return isNaN(n) ? 0 : n;
  }

  function getIdx(map, name) {
    return map.hasOwnProperty(name) ? map[name] : -1;
  }

  function setRowValue(arr, idx, val) {
    if (idx >= 0 && idx < arr.length) {
      arr[idx] = val;
    }
  }

  function getRowValue(arr, idx) {
    if (idx >= 0 && idx < arr.length) return arr[idx];
    return '';
  }

  function getInventoryBalanceMap() {
    var resultMap = {};
    
    var inventorybalanceSearchObj = search.create({
      type: 'inventorybalance',
      filters: [
        ['status', 'anyof', '6', '1'],
        'AND',
        ['item.isinactive', 'is', 'F']
      ],
      columns: [
        search.createColumn({
          name: 'item',
          summary: 'GROUP'
        }),
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
          name: 'available',
          summary: 'SUM',
          label: 'Available'
        })
      ]
    });
    
    inventorybalanceSearchObj.run().each(function(result) {
      var itemId = result.getValue({ name: 'item', summary: 'GROUP' });
      var goodQty = parseFloat(result.getValue({ name: 'formulanumeric', summary: 'SUM' })) || 0;
      var badQty  = parseFloat(result.getValue({ name: 'formulanumeric1', summary: 'SUM' })) || 0;
      var total = parseFloat((goodQty + badQty).toFixed(2));
      var avail = parseFloat(result.getValue({ name: 'available', summary: 'SUM' })) || 0;
      
      resultMap[itemId] = {
        good: goodQty,
        bad: badQty,
        total: total,
        avail: avail
      };
      return true;
    });
    
    return resultMap;
  }

  function getBilledExpiryMap(fileId) {
    var resultMap = {};
    if (!fileId) return resultMap;

    var f = file.load({ id: fileId });
    var rows = f.getContents().split('\n');

    var today = new Date();
    today = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    for (var i = 1; i < rows.length; i++) {
      var line = rows[i];
      if (!line) continue;

      var cols = splitCsvRow(line);

      var itemId = (cols[0] || '').replace(/"/g, '').trim();
      var maxDate = (cols[1] || '').replace(/"/g, '').trim();
      var shelfLife = parseInt(cols[2], 10) || 0;

      if (!itemId || !maxDate) continue;

      if (!shelfLife || shelfLife === 0) {
        resultMap[itemId] = {
          billedDate: maxDate,
          expireDate: '',
          stat: ''
        };
        continue;
      }

      var d = maxDate.split('/');
      var billedDate = new Date(d[2], d[0] - 1, d[1]);

      var expiryDate = new Date(billedDate);
      expiryDate.setDate(expiryDate.getDate() + shelfLife);

      var diffDays = Math.floor((expiryDate - today) / (1000 * 60 * 60 * 24));

      var status = 'OK';

      if (today > expiryDate) {
        status = 'Expired';
      } else if (diffDays <= 120) {
        status = 'Expiring';
      }

      resultMap[itemId] = {
        billedDate: maxDate,
        expireDate: (expiryDate.getMonth() + 1) + '/' + expiryDate.getDate() + '/' + expiryDate.getFullYear(),
        stat: status
      };
    }

    return resultMap;
  }

  function generateCsvFile(isCron) {
    var fileIdItem     = null;
    var fileIdInv      = null;
    var fileIdAssembly = null;
    var fileIdBilled   = null;
    
    var folderSearch = search.create({
      type: 'folder',
      filters: [
        ['internalid', 'anyof',
          String(FOLDER_ITEM_LIST),
          String(FOLDER_INV_ITEM_LIST),
          String(FOLDER_ASSEMBLY),
          String(lastBilledFile)
        ],
        'AND',
        ['file.documentsize', 'greaterthan', '3']
      ],
      columns: [
        search.createColumn({ name: 'internalid', summary: 'GROUP' }),
        search.createColumn({ name: 'internalid', join: 'file', summary: 'MAX' })
      ]
    });
    
    folderSearch.run().each(function (result) {
      var folderId = parseInt(result.getValue({ name: 'internalid', summary: 'GROUP' }), 10);
      var fId = result.getValue({ name: 'internalid', join: 'file', summary: 'MAX' });
      
      if (folderId === FOLDER_ITEM_LIST) {
        fileIdItem = fId;
      } else if (folderId === FOLDER_INV_ITEM_LIST) {
        fileIdInv = fId;
      } else if (folderId === FOLDER_ASSEMBLY) {
        fileIdAssembly = fId;
      } else if (folderId === lastBilledFile) {
        fileIdBilled = fId;
      }
      return true;
    });
    
    if (!fileIdItem && !fileIdInv && !fileIdAssembly && !fileIdBilled) {
      throw new Error('No files found in the specified folders.');
    }
    
    var rowsItem     = loadCsvRows(fileIdItem);
    var rowsInv      = loadCsvRows(fileIdInv);
    var rowsAssembly = loadCsvRows(fileIdAssembly);
    var rowsBilled   = getBilledExpiryMap(fileIdBilled);

    log.debug('fileIdBilled', fileIdBilled);
    log.debug('rowsBilled', rowsBilled);
    
    if (!rowsItem || rowsItem.length === 0) {
      throw new Error('Item list file is empty or missing header.');
    }
    
    var headerRow = rowsAssembly[0];
    var allRows = [];
    allRows = allRows.concat(rowsAssembly);
    if (rowsInv.length > 1) {
      allRows = allRows.concat(rowsInv.slice(1));
    }
    if (rowsItem.length > 1) {
      allRows = allRows.concat(rowsItem.slice(1));
    }
    
    var newContent = [];
    var headers = splitCsvRow(headerRow);
    var headersClean = headers.map(cleanHeader);

    var headerMap = {};
    for (var h = 0; h < headersClean.length; h++) {
      headerMap[headersClean[h]] = h;
    }

    var IDX_ITEM_ID                         = getIdx(headerMap, 'Item ID');
    var IDX_WARNING_LT_2                    = getIdx(headerMap, 'Warning (< 2 Months)');
    var IDX_RECOMMENDED_RESTRICTION_QTY     = getIdx(headerMap, 'Recommened Restriction Quantity');
    var IDX_MAX_COMMITTED                   = getIdx(headerMap, 'Maximum of Committed');
    var IDX_ON_HOLD                         = getIdx(headerMap, 'Sum of On Hold');
    var IDX_RESERVED                        = getIdx(headerMap, 'Sum of Reserved');
    var IDX_ON_HAND_AVAIL_GOOD              = getIdx(headerMap, 'Sum of On Hand Available (Good)');
    var IDX_ON_HAND_TOTAL                   = getIdx(headerMap, 'Sum of On Hand Total (Good + Deviated)');
    var IDX_IN_TRANSIT                      = getIdx(headerMap, 'In Transit');
    var IDX_ON_HAND_TOTAL_TRANSIT           = getIdx(headerMap, 'Sum of On Hand Total + In Transit');
    var IDX_ON_ORDER                        = getIdx(headerMap, 'On Order (TO SHIP)');
    var IDX_TOTAL_STOCK                     = getIdx(headerMap, 'Total Stock');
    var IDX_NEXT_ARRIVAL_QTY                = getIdx(headerMap, 'Next Arrival Qty');
    var IDX_NEXT_ARRIVAL_DATE               = getIdx(headerMap, 'Next Arrival Date');
    var IDX_DAYS_TILL_NEXT_ARRIVAL          = getIdx(headerMap, 'Days Till Next Arrival');
    var IDX_MONTHS_TILL_NEXT_ARRIVAL        = getIdx(headerMap, 'Months Till Next Arrival');
    var IDX_ON_HAND_TOTAL_MONTHS            = getIdx(headerMap, 'On Hand Total (Months)');
    var IDX_INSPECTION                      = getIdx(headerMap, 'Inspection');
    var IDX_LABEL                           = getIdx(headerMap, 'Label');
    var IDX_DEVIATION                       = getIdx(headerMap, 'Deviation');
    var IDX_ON_HAND_TRANSIT_MONTHS          = getIdx(headerMap, 'On Hand Total + Transit (Months)');
    var IDX_TOTAL_STOCK_MONTHS              = getIdx(headerMap, 'Total Stock [Months]');
    var IDX_30_DAYS                         = getIdx(headerMap, '30 Days');
    var IDX_60_DAYS                         = getIdx(headerMap, '60 Days');
    var IDX_90_DAYS                         = getIdx(headerMap, '90 Days');
    var IDX_120_DAYS                        = getIdx(headerMap, '120 Days');
    var IDX_4_MONTH_AVG                     = getIdx(headerMap, '4 Month Average');
    var IDX_SHELF_LIFE                      = getIdx(headerMap, 'Maximum of Shelf Life in Days');

    headersClean.push('Last Billed Date');
    headersClean.push('Expire Date');
    headersClean.push('Expiry Status');
    newContent.push(headersClean);
    
    var seenItemIds = {};
    var rowObjs = [];
    
    var balances = getInventoryBalanceMap();
    log.debug('Inventory Balance Map', balances);
    
    allRows.slice(1).forEach(function (line) {
      if (!line || line.trim() === '') return;
      
      var colsRaw = splitCsvRow(line);
      var rawItemId = (colsRaw[IDX_ITEM_ID] || '').replace(/"/g, '').trim();
      
      if (!rawItemId) return;
      if (seenItemIds[rawItemId]) return;
      seenItemIds[rawItemId] = true;
      
      var cleaned = colsRaw.map(function (v) { return cleanCsvValue(v); });
      var itemIdNum = parseInt(rawItemId, 10);
      if (!isFinite(itemIdNum)) itemIdNum = 0;
      
      rowObjs.push({
        itemIdNum: itemIdNum,
        cleaned: cleaned
      });
    });
    
    rowObjs.sort(function (a, b) {
      return a.itemIdNum - b.itemIdNum;
    });

    var displayRows = [];
    
    rowObjs.forEach(function (obj) {
      var cleaned = obj.cleaned;
      var displayRow = [];
      var itemid = getRowValue(cleaned, IDX_ITEM_ID);

      var good = 0;
      var bad = 0;
      var total = 0;
      var avail = 0;

      if (balances[itemid]) {
        good  = balances[itemid].good;
        bad   = balances[itemid].bad;
        total = balances[itemid].total;
        avail = balances[itemid].avail;
      }

      var inTransit = toNumber(getRowValue(cleaned, IDX_IN_TRANSIT));
      var onOrder   = toNumber(getRowValue(cleaned, IDX_ON_ORDER)) - inTransit;
      var avg       = toNumber(getRowValue(cleaned, IDX_4_MONTH_AVG)) / 4;
      var daysTillAvail = toNumber(getRowValue(cleaned, IDX_DAYS_TILL_NEXT_ARRIVAL));
      var onHandMonth = 0;

      setRowValue(cleaned, IDX_ON_HOLD, bad);
      setRowValue(cleaned, IDX_ON_HAND_AVAIL_GOOD, good);
      setRowValue(cleaned, IDX_ON_HAND_TOTAL, total);

      if (IDX_ON_HAND_TOTAL_TRANSIT >= 0) {
        setRowValue(cleaned, IDX_ON_HAND_TOTAL_TRANSIT, parseFloat(good) + parseFloat(inTransit || 0));
      }

      if (IDX_ON_ORDER >= 0) {
        setRowValue(cleaned, IDX_ON_ORDER, onOrder);
      }

      if (IDX_TOTAL_STOCK >= 0) {
        var qtyTransit = toNumber(getRowValue(cleaned, IDX_ON_HAND_TOTAL_TRANSIT));
        var qtyOnOrder = toNumber(getRowValue(cleaned, IDX_ON_ORDER));
        setRowValue(cleaned, IDX_TOTAL_STOCK, parseFloat(qtyTransit) + parseFloat(qtyOnOrder));
      }

      if (IDX_ON_HAND_TOTAL_MONTHS >= 0) {
        if (avg === 0) {
          setRowValue(cleaned, IDX_ON_HAND_TOTAL_MONTHS, '');
        } else {
          onHandMonth = parseFloat((parseFloat(good) / avg).toFixed(2));
          setRowValue(cleaned, IDX_ON_HAND_TOTAL_MONTHS, onHandMonth);
        }
      }

      if (IDX_ON_HAND_TRANSIT_MONTHS >= 0) {
        if (avg === 0) {
          setRowValue(cleaned, IDX_ON_HAND_TRANSIT_MONTHS, '');
        } else {
          setRowValue(cleaned, IDX_ON_HAND_TRANSIT_MONTHS, ((parseFloat(good) + parseFloat(inTransit || 0)) / avg).toFixed(2));
        }
      }

      if (IDX_TOTAL_STOCK_MONTHS >= 0) {
        if (avg === 0) {
          setRowValue(cleaned, IDX_TOTAL_STOCK_MONTHS, '');
        } else {
          setRowValue(cleaned, IDX_TOTAL_STOCK_MONTHS, ((parseFloat(avail)) / avg).toFixed(2));
        }
      }

      if (IDX_WARNING_LT_2 >= 0) {
        var yorn = 'No';
        if (onHandMonth <= 2 && (daysTillAvail === 0 || daysTillAvail > 30)) yorn = 'Yes';
        setRowValue(cleaned, IDX_WARNING_LT_2, yorn);
      }

      // Preserve core logic area for editable/input column
      // old UI used cIdx === 9 as editable field, which now maps to:
      // "Recommened Restriction Quantity"
      // not changing business logic, only mapping.
      
      // Build display row after all overrides
      for (var cIdx = 0; cIdx < cleaned.length; cIdx++) {
        var value = cleaned[cIdx];
        var txt = String(value || '').replace(/^"+|"+$/g, '');
        displayRow.push(txt);
      }

      if (rowsBilled && rowsBilled[itemid]) {
        var relatedDate = rowsBilled[itemid];
        log.audit('relatedDate', relatedDate);
        displayRow.push(relatedDate.billedDate);
        displayRow.push(relatedDate.expireDate);
        displayRow.push(relatedDate.stat);
        cleaned.push(relatedDate.billedDate);
        cleaned.push(relatedDate.expireDate);
        cleaned.push(relatedDate.stat);
      } else {
        displayRow.push('');
        displayRow.push('');
        displayRow.push('');
        cleaned.push('');
        cleaned.push('');
        cleaned.push('');
      }
    
      newContent.push(cleaned);
      displayRows.push(displayRow);
    });
    
    var cleanedCsv = newContent.map(function (row) {
      return row.join(',');
    }).join('\n');

    var targetFolder = isCron ? FOLDER_DOWNLOAD_CRON : FOLDER_DOWNLOAD_UI;
    var fileName = isCron
      ? 'Item Avail Tool.csv'
      : 'Avail_Download_' + (new Date().getTime()) + '.csv';
    
    var newFileObj = file.create({
      name: fileName,
      fileType: file.Type.CSV,
      contents: cleanedCsv,
      encoding: file.Encoding.UTF8,
      folder: targetFolder,
      isOnline: true
    });
    var newFileId = newFileObj.save();
    
    var reloadFile = file.load({ id: newFileId });
    var dlUrl = reloadFile.url || '';

    return {
      fileId: newFileId,
      url: dlUrl,
      name: reloadFile.name || '',
      headersClean: headersClean,
      displayRows: displayRows,
      newContent: newContent,
      cleanedCsv: cleanedCsv,
      recommendedRestrictionIdx: IDX_RECOMMENDED_RESTRICTION_QTY,
      warningIdx: IDX_WARNING_LT_2
    };
  }
  
  function onRequest(context) {
    log.debug('Triggered');
    if (context.request.method !== 'GET') {
      context.response.write('This Suitelet only supports GET.');
      return;
    }
    
    var q = context.request.parameters || {};
    var mode    = q.mode || '';
    var cronts  = q.cronts || '';
    var cronsig = q.cronsig || '';

    if (mode === 'cron') {
      context.response.setHeader({
        name: 'Content-Type',
        value: 'application/json'
      });

      if (!verifyCron(cronts, cronsig)) {
        context.response.write(JSON.stringify({
          ok: false,
          message: 'Invalid cron token'
        }));
        return;
      }

      try {
        var cronResult = generateCsvFile(true);
        context.response.write(JSON.stringify({
          ok: true,
          fileId: cronResult.fileId,
          url: cronResult.url,
          name: cronResult.name
        }));
      } catch (e) {
        log.error('cron mode error', e);
        context.response.write(JSON.stringify({
          ok: false,
          message: e.message || String(e)
        }));
      }
      return;
    }

    var form = ui.createForm({ title: 'Availability Tool' });
    
    var empid = q.empid || '';
    var ts    = q.ts    || '';
    var sig   = q.sig   || '';
    
    var selectedEmp = q.custpage_id || '';
    
    if (empid && ts && sig && verify(empid, ts, sig)) {
      selectedEmp = empid;
    }
    
    if (!selectedEmp) {
      context.response.write(
        '<html><head>' +
        '<script>setTimeout(function(){ window.location.href = ' + JSON.stringify(PORTAL_URL) + '; }, 1200);</script>' +
        '<style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>' +
        '</head><body><div class="message">Login Required</div></body></html>'
      );
      return;
    }
    
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
      label: 'Availability Table'
    });

    var result;
    try {
      result = generateCsvFile(false);
    } catch (e) {
      log.error('generateCsvFile error', e);
      htmlField.defaultValue = '<p style="color:red;">' + (e.message || e.toString()) + '</p>';
      context.response.writePage(form);
      return;
    }

    fileField.defaultValue = result.fileId;
    
    var html = '';
    html += '<style>' +
      '.table-container{max-height:850px;overflow-y:auto;overflow-x:auto;border:1px solid #ccc;}' +
      '.h-scroll{height:16px;overflow-x:auto;overflow-y:hidden;border:1px solid #ccc;border-bottom:0;width:100%;}' +
      '.h-scroll-inner{height:1px;}' +
      '#excelTable{border-collapse:separate;width:max-content;min-width:100%;table-layout:auto;}' +
      '#excelTable th,#excelTable td{border:1px solid #ccc;padding:6px 10px;background-color:#fff;white-space:nowrap;font-size:12px;}' +
      '#excelTable thead th{position:sticky;top:0;background-color:#f3f3f3;z-index:9;}' +
      
      '.download-btn{background-color:white;border:none;font-size:13px;cursor:pointer;margin-bottom:8px;}' +
      '.download-btn:hover{background-color:#eef7ff;}' +

      '.row-warning td{background-color:#ffd6d6 !important;}' +
      '.row-warning .sticky-col{background-color:#ffd6d6 !important;}' +

      '.row-expired td{background-color:#ead6ff !important;}' +
      '.row-expired .sticky-col{background-color:#ead6ff !important;}' +

      '.row-expiring td{background-color:#fff7bf !important;}' +
      '.row-expiring .sticky-col{background-color:#fff7bf !important;}' +

      '.toolbar-wrap{display:flex;align-items:center;justify-content:flex-start;gap:12px;flex-wrap:wrap;margin-bottom:8px;}' +
      '.legend-wrap{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}' +
      '.legend-label{font-size:12px;font-weight:600;color:#333;}' +
      '.legend-pill{display:inline-flex;align-items:center;gap:6px;padding:5px 10px;border:1px solid #ccc;border-radius:999px;cursor:pointer;background:#fff;font-size:12px;user-select:none;}' +
      '.legend-pill:hover{background:#f8f8f8;}' +
      '.legend-pill.active{border-color:#2563eb;box-shadow:0 0 0 2px rgba(37,99,235,.12);}' +
      '.legend-dot{width:12px;height:12px;border-radius:50%;display:inline-block;border:1px solid rgba(0,0,0,.15);}' +
      '.legend-red .legend-dot{background:#ffd6d6;}' +
      '.legend-purple .legend-dot{background:#ead6ff;}' +
      '.legend-yellow .legend-dot{background:#fff7bf;}' +
      
      '.th-filter-wrap{position:relative;display:inline-flex;align-items:center;gap:6px;}' +
      '.th-filter-btn{cursor:pointer;border:1px solid #cbd5e1;background:#fff;padding:2px 4px;border-radius:4px;line-height:1;font-size:11px;}' +
      '.th-filter-btn:hover{background:#f3f4f6;}' +
      '.th-filter-panel{position:fixed;top:0;left:0;width:260px;max-height:320px;overflow:auto;background:#fff;border:1px solid #cbd5e1;border-radius:8px;padding:10px;box-shadow:0 8px 24px rgba(0,0,0,.12);z-index:9999;}' +
      '.th-filter-panel .hdr{font-weight:600;font-size:12px;margin-bottom:6px;}' +
      '.th-filter-panel input[type="search"]{width:100%;padding:6px 8px;box-sizing:border-box;margin-bottom:8px;}' +
      '.th-filter-panel .list{max-height:200px;overflow:auto;border:1px solid #e5e7eb;border-radius:6px;padding:6px;}' +
      '.th-filter-panel .row{display:flex;align-items:center;gap:8px;padding:2px 0;font-size:12px;}' +
      '.th-filter-panel .actions{display:flex;justify-content:space-between;gap:8px;margin-top:8px;}' +
      '.th-filter-panel .actions button{padding:2px 6px;font-size:11px;line-height:1.2;border-radius:4px;border:1px solid #cbd5e1;background:#fff;cursor:pointer;}' +
      '.th-filter-panel .actions button[data-act="apply"]{font-weight:600;border-color:#9ab3ff;}' +
      '.th-filter-panel .actions button[data-act="clear"]{color:#7f1d1d;border-color:#f3d2d2;}' +
      '.th-filter-active .th-filter-btn{border-color:#2563eb;background:#eff6ff;}' +
      
      '.sticky-col{position:sticky;background:#f9f9f9;background-clip:padding-box;z-index:11;}' +
      'thead th.sticky-col{z-index:30;}' +
      '.sticky-col.sep-left{border-left-color:transparent;box-shadow:inset 1px 0 0 #ccc;}' +
      '.col-sticky-0{left:0px;min-width:180px;}' +
      '.col-sticky-1{left:180px;min-width:110px;}' +
      '.col-sticky-2{left:290px;min-width:180px;}' +
      '.col-sticky-3{left:470px;min-width:160px;}' +
      '</style>';
    
    html += '<div class="toolbar-wrap">' +
      '<button id="downloadCsvBtn" class="download-btn" type="button">' +
      '<img src="https://cdn-icons-png.flaticon.com/512/10630/10630240.png" ' +
      'alt="csv-icon" style="width:16px;vertical-align:middle;margin-right:6px;" />' +
      'Download CSV</button>' +

      '<div class="legend-wrap" id="colorLegendWrap">' +
        '<span class="legend-label">Color Legend:</span>' +
        '<span class="legend-pill legend-red" data-row-filter="warning">' +
          '<span class="legend-dot"></span><span>Less than 2 Months</span>' +
        '</span>' +
        '<span class="legend-pill legend-purple" data-row-filter="expired">' +
          '<span class="legend-dot"></span><span>Expired</span>' +
        '</span>' +
        '<span class="legend-pill legend-yellow" data-row-filter="expiring">' +
          '<span class="legend-dot"></span><span>Expiring</span>' +
        '</span>' +
      '</div>' +
    '</div>';

    html += '<div id="filter-portal"></div>';
    html += '<div class="h-scroll" id="topScroll"><div class="h-scroll-inner" id="topScrollInner"></div></div>';
    html += '<div class="table-container"><table id="excelTable"><thead><tr>';
    
    result.headersClean.forEach(function (hVal, idx) {
      var label = hVal || ('Col ' + (idx + 1));
      html += '<th data-col-idx="' + idx + '">' +
        '<span class="th-filter-wrap">' +
        '<span class="th-label">' + label + '</span>' +
        '<button class="th-filter-btn" type="button" data-col="' + idx + '">▾</button>' +
        '</span></th>';
    });
    
    html += '</tr></thead><tbody>';

    var expiryStatusIndex = result.headersClean.indexOf('Expiry Status');
    var warningColIndex = result.headersClean.indexOf('Warning (< 2 Months)');
    var editableColIndex = result.headersClean.indexOf('Recommened Restriction Quantity');

    result.displayRows.forEach(function (row) {
      var rowClass = '';
      var warningLessThan2 = warningColIndex >= 0 ? String(row[warningColIndex] || '').trim().toLowerCase() : '';
      var expiryStatus = String(row[expiryStatusIndex] || '').trim().toLowerCase();

      if (warningLessThan2 === 'yes') {
        rowClass = ' class="row-warning"';
      } else if (expiryStatus === 'expired') {
        rowClass = ' class="row-expired"';
      } else if (expiryStatus === 'expiring') {
        rowClass = ' class="row-expiring"';
      }

      html += '<tr' + rowClass + '>';

      for (var cIdx = 0; cIdx < row.length; cIdx++) {
        var txt = row[cIdx];

        if (cIdx === editableColIndex) {
          html += '<td style="width:80px;min-width:80px;max-width:80px;">' +
            '<input type="number" value="' + String(txt || '').replace(/"/g, '') +
            '" style="width:100%;box-sizing:border-box;" />' +
            '</td>';
        } else {
          html += '<td>' + txt + '</td>';
        }
      }

      html += '</tr>';
    });

    html += '</tbody></table></div>';
    html += '<script>window.DOWNLOAD_URL = ' + JSON.stringify(result.url) + ';</script>';
    
    html += '<script>' +
      'document.addEventListener("DOMContentLoaded", function(){' +
      'var exportBtn = document.getElementById("downloadCsvBtn");' +
      'if(exportBtn){exportBtn.addEventListener("click", function(e){e.preventDefault();e.stopPropagation();var url=(window.DOWNLOAD_URL||"").trim();if(url){window.open(url,"_blank");}else{alert("Download not available yet.");}});}' +
      
      'var table = document.getElementById("excelTable");' +
      'var container = document.querySelector(".table-container");' +
      'var topScroll = document.getElementById("topScroll");' +
      'var topInner = document.getElementById("topScrollInner");' +
      'var warningColIndex = -1;' +
      'var editableColIndex = -1;' +
      
      'function updateTopScrollbarWidth(){' +
      'if(!table||!topInner||!container) return;' +
      'var w = Math.max(table.scrollWidth||0, (container.clientWidth||0)+2);' +
      'topInner.style.width = w + "px";' +
      '}' +
      'if(topScroll && container){' +
      'topScroll.addEventListener("scroll", function(){container.scrollLeft = topScroll.scrollLeft;});' +
      'container.addEventListener("scroll", function(){topScroll.scrollLeft = container.scrollLeft;});' +
      '}' +
      'updateTopScrollbarWidth();' +
      'window.addEventListener("resize", updateTopScrollbarWidth);' +
      
      'var activeFilters = {};' +
      'var activeLegendFilters = new Set();' +
      'var expiryStatusColIdx = -1;' +
      'for (var i = 0; i < table.tHead.rows[0].cells.length; i++) {' +
      '  var labelNode = table.tHead.rows[0].cells[i].querySelector(".th-label");' +
      '  var hdr = labelNode ? (labelNode.textContent || "").trim().toLowerCase() : "";' +
      '  if (hdr === "expiry status") expiryStatusColIdx = i;' +
      '  if (hdr === "warning (< 2 months)") warningColIndex = i;' +
      '  if (hdr === "recommened restriction quantity") editableColIndex = i;' +
      '}' +

      'function bodyRows(){return table && table.tBodies[0] ? table.tBodies[0].rows : [];}' +
      'function getCellText(row, idx){var cells=row.cells;if(!cells||idx<0||idx>=cells.length) return "";var t="";if(cells[idx].querySelector("input")){t=cells[idx].querySelector("input").value||"";}else{t=cells[idx].textContent||"";}return String(t).trim();}' +
      'function normalizeVal(v){var s=String(v==null?"":v).trim();if(/^-+\\s*none\\s*-+$/i.test(s))s="";if(/^nan$/i.test(s))s="";return s.toLowerCase();}' +
      
      'function getAllValuesForColumn(colIdx){' +
      'var rows=bodyRows();var displayByKey={};' +
      'for(var r=0;r<rows.length;r++){' +
      'var v=getCellText(rows[r], colIdx);' +
      'if(v===" - None -"||v==="NaN") v="";' +
      'if(v===".00") v="0.00";' +
      'var key=normalizeVal(v);' +
      'if(!displayByKey[key]) displayByKey[key]=(key===""?"(blank)":(v||"(blank)"));' +
      '}' +
      'var keys = Object.keys(displayByKey).sort(function(a,b){var da=displayByKey[a], db=displayByKey[b];if(a===""&&b==="")return 0;if(a==="")return 1;if(b==="")return -1;return da.localeCompare(db,undefined,{numeric:true,sensitivity:"base"});});' +
      'return {keys:keys,displayByKey:displayByKey};' +
      '}' +
      
      'function rowPassesFiltersExceptColumn(row, exceptIdx){' +
      'for(var k in activeFilters){if(!Object.prototype.hasOwnProperty.call(activeFilters,k))continue;var idx=parseInt(k,10);if(idx===exceptIdx)continue;var set=activeFilters[k];if(set&&set.size>0){var v=getCellText(row,idx).toLowerCase();if(!set.has(v))return false;}}return true;' +
      '}' +
      
      'function getValueCountsForColumn(colIdx){' +
      'var rows=bodyRows(),counts={};' +
      'for(var r=0;r<rows.length;r++){' +
      'var row=rows[r];' +
      'if(!rowPassesFiltersExceptColumn(row,colIdx)) continue;' +
      'var v=getCellText(rows[r], colIdx);' +
      'if(v===" - None -"||v==="NaN") v="";' +
      'if(v===".00") v="0.00";' +
      'var key=normalizeVal(v);' +
      'counts[key]=(counts[key]||0)+1;' +
      '}' +
      'return counts;' +
      '}' +
      
      'function applyFilters(){' +
      '  var rows = bodyRows(), pairs = [];' +
      '  for (var k in activeFilters) {' +
      '    if (!Object.prototype.hasOwnProperty.call(activeFilters, k)) continue;' +
      '    var s = activeFilters[k];' +
      '    if (s && s.size > 0) pairs.push([parseInt(k, 10), s]);' +
      '  }' +
      '  for (var r = 0; r < rows.length; r++) {' +
      '    var row = rows[r], show = true;' +
      '    for (var i = 0; i < pairs.length && show; i++) {' +
      '      var colIdx = pairs[i][0], set = pairs[i][1];' +
      '      var val = getCellText(row, colIdx).toLowerCase();' +
      '      if (!set.has(val)) show = false;' +
      '    }' +
      '    if (show && activeLegendFilters.size > 0) {' +
      '      var matched = false;' +
      '      var lessThan2Val = warningColIndex >= 0 ? getCellText(row, warningColIndex).toLowerCase() : "";' +
      '      var expiryVal = expiryStatusColIdx >= 0 ? getCellText(row, expiryStatusColIdx).toLowerCase() : "";' +
      '      activeLegendFilters.forEach(function(key){' +
      '        if (key === "warning" && lessThan2Val === "yes") matched = true;' +
      '        if (key === "expired" && expiryVal === "expired") matched = true;' +
      '        if (key === "expiring" && expiryVal === "expiring") matched = true;' +
      '      });' +
      '      if (!matched) show = false;' +
      '    }' +
      '    row.style.display = show ? "" : "none";' +
      '  }' +
      '}' +

      'var legendWrap = document.getElementById("colorLegendWrap");' +
      'if(legendWrap){' +
      '  legendWrap.addEventListener("click", function(e){' +
      '    var pill = e.target.closest(".legend-pill");' +
      '    if(!pill) return;' +
      '    var key = pill.getAttribute("data-row-filter") || "";' +
      '    if(!key) return;' +
      '    if(activeLegendFilters.has(key)){' +
      '      activeLegendFilters.delete(key);' +
      '      pill.classList.remove("active");' +
      '    } else {' +
      '      activeLegendFilters.add(key);' +
      '      pill.classList.add("active");' +
      '    }' +
      '    applyFilters();' +
      '  });' +
      '}' +
      
      'function openFilterPanel(btn,colIdx){' +
      'var th=btn.closest("th");' +
      'var oldPanel=document.querySelector("#filter-portal .th-filter-panel");' +
      'if(oldPanel){var same=(oldPanel.__ownerTH===th);oldPanel.remove();if(th)th.classList.remove("th-filter-active");if(same)return;}' +
      'document.querySelectorAll("th.th-filter-active").forEach(function(node){node.classList.remove("th-filter-active");});' +
      
      'var panel=document.createElement("div");' +
      'panel.className="th-filter-panel";' +
      'panel.innerHTML=' +
      '"<div class=\\"hdr\\">Filter</div>" +' +
      '"<input type=\\"search\\" placeholder=\\"Search values...\\">" +' +
      '"<div class=\\"list\\"></div>" +' +
      '"<div class=\\"actions\\">" +' +
      '"<button type=\\"button\\" data-act=\\"clear\\">Clear</button>" +' +
      '"<div style=\\"display:flex;gap:6px;\\">" +' +
      '"<button type=\\"button\\" data-act=\\"selectall\\">Select all</button>" +' +
      '"<button type=\\"button\\" data-act=\\"deselectall\\">Deselect all</button>" +' +
      '"<button type=\\"button\\" data-act=\\"apply\\">Apply</button>" +' +
      '"</div>" +' +
      '"</div>";' +
      
      'var list=panel.querySelector(".list");' +
      'panel.__ownerTH = th;' +
      
      'var grouped=getAllValuesForColumn(colIdx);' +
      'var universeKeys=grouped.keys;' +
      'var displayByKey=grouped.displayByKey;' +
      'var visibleCounts=getValueCountsForColumn(colIdx);' +
      'var existing=activeFilters[colIdx]||null;' +
      'var tempSel=new Set();' +
      
      'if(existing && existing.size>0){' +
      'universeKeys.forEach(function(k){if(existing.has(k)) tempSel.add(k);});' +
      '}else{' +
      'universeKeys.forEach(function(k){if((visibleCounts[k]||0)>0) tempSel.add(k);});' +
      '}' +
      
      'function renderList(filterText){' +
      'list.innerHTML="";' +
      'var ft=String(filterText||"").toLowerCase();' +
      'universeKeys.forEach(function(k){' +
      'var labelText=displayByKey[k]||(k===""?"(blank)":k);' +
      'if(ft && labelText.toLowerCase().indexOf(ft)===-1) return;' +
      'var id="f_"+colIdx+"_"+Math.random().toString(36).slice(2);' +
      'var checked=tempSel.has(k);' +
      'var cnt=visibleCounts[k]||0;' +
      'var row=document.createElement("div");row.className="row";row.dataset.key=k;' +
      'row.innerHTML="<input type=\\"checkbox\\" id=\\""+id+"\\" "+(checked?"checked":"")+">" +"<label for=\\""+id+"\\">"+labelText+(cnt?"  ("+cnt+")":"")+"</label>";' +
      'list.appendChild(row);' +
      '});' +
      '}' +
      'renderList("");' +
      
      'panel.querySelector("input[type=search]").addEventListener("input", function(){renderList(this.value);});' +
      
      'list.addEventListener("change", function(e){var cb=e.target;if(!cb||cb.type!=="checkbox")return;var row=cb.closest(".row");if(!row)return;var key=row.dataset.key||"";if(cb.checked)tempSel.add(key);else tempSel.delete(key);});' +
      
      'panel.querySelector(".actions").addEventListener("click", function(e){var act=e.target && e.target.getAttribute("data-act");if(!act)return;' +
      'if(act==="clear"){delete activeFilters[colIdx];if(th)th.classList.remove("th-filter-active");panel.remove();applyFilters();return;}' +
      'if(act==="selectall"){var cbs=list.querySelectorAll("input[type=checkbox]");for(var i=0;i<cbs.length;i++){var cb=cbs[i];if(!cb.checked)cb.checked=true;var key=cb.closest(".row").dataset.key||"";tempSel.add(key);}return;}' +
      'if(act==="deselectall"){var cbs2=list.querySelectorAll("input[type=checkbox]");for(var j=0;j<cbs2.length;j++){var cb2=cbs2[j];if(cb2.checked)cb2.checked=false;var key2=cb2.closest(".row").dataset.key||"";tempSel.delete(key2);}return;}' +
      'if(act==="apply"){' +
      'var allow=new Set();tempSel.forEach(function(k){allow.add(k);});' +
      'if(allow.size===universeKeys.length){delete activeFilters[colIdx];if(th)th.classList.remove("th-filter-active");}' +
      'else{activeFilters[colIdx]=allow;if(th)th.classList.add("th-filter-active");}' +
      'panel.remove();applyFilters();}' +
      '});' +
      
      'var portal=document.getElementById("filter-portal") || document.body;' +
      'portal.appendChild(panel);' +
      
      'function positionPanel(){' +
      'var rect=btn.getBoundingClientRect();' +
      'var vw=Math.max(document.documentElement.clientWidth||0, window.innerWidth||0);' +
      'var ph=panel.offsetHeight||0;var pw=panel.offsetWidth||260;' +
      'var belowTop=rect.bottom+6;var aboveTop=rect.top-ph-6;' +
      'var top=(belowTop+ph<=window.innerHeight)?belowTop:Math.max(8,aboveTop);' +
      'var left=Math.min(vw-pw-8, Math.max(8, rect.right-pw));' +
      'panel.style.top=Math.round(top)+"px";panel.style.left=Math.round(left)+"px";' +
      '}' +
      'positionPanel();' +
      'var onScrollResize=function(){positionPanel();};' +
      'window.addEventListener("scroll", onScrollResize, true);' +
      'window.addEventListener("resize", onScrollResize);' +
      'if(container) container.addEventListener("scroll", onScrollResize);' +
      
      'setTimeout(function(){document.addEventListener("click", function handler(e){if(panel.contains(e.target)||th.contains(e.target))return;panel.remove();if(th)th.classList.remove("th-filter-active");window.removeEventListener("scroll",onScrollResize,true);window.removeEventListener("resize",onScrollResize);if(container)container.removeEventListener("scroll",onScrollResize);document.removeEventListener("click",handler);});},0);' +
      
      'if(th)th.classList.add("th-filter-active");' +
      '}' +
      
      'if(table && table.tHead){' +
      'table.tHead.addEventListener("click", function(e){var btn=e.target.closest(".th-filter-btn");if(!btn)return;var colIdx=parseInt(btn.getAttribute("data-col"),10);if(!isFinite(colIdx))return;e.stopPropagation();openFilterPanel(btn,colIdx);});' +
      '}' +
      '});' +
      '</script>';
    
    htmlField.defaultValue = html;
    context.response.writePage(form);
  }

  return {
    onRequest: onRequest
  };
});