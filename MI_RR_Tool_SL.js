/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/ui/serverWidget', 'N/file', 'N/log', 'N/search', 'N/record', 'N/runtime', 'N/crypto'],
  function (ui, file, log, search, record, runtime, crypto) {
    function onRequest(context) {

      var portalUrl = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2110&deploy=1&compid=4975346&ns-at=AAEJ7tMQamzukv1WMqTK6i2c27bRetbrd2MDLjhDgPPFOawMxCo';

      // --- Signed session helpers ---
      const SECRET = runtime.getCurrentScript().getParameter({ name: 'custscript_portal_secret' }) || 'change-me';
      const TOKEN_TTL_MS = 30 * 60 * 1000; // 30 minutes

      function signCron(ts) {
        var h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
        h.update({ input: 'CRON|' + ts + '|rR9Z7KpXw2N6C8mE4HqFJYvT5bS0aUeD1LQG3oM' });
        return h.digest({ outputEncoding: crypto.Encoding.HEX });
      }

      function verifyCron(ts, sig) {
        if (!ts || !sig) return false;
        if (Math.abs(Date.now() - parseInt(ts, 10)) > TOKEN_TTL_MS) return false;
        try { return signCron(ts) === sig; } catch (e) { log.error('verifyCron token', e); return false; }
      }

      function sign(empid, ts) {
        var h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
        h.update({ input: empid + '|' + ts + '|' + SECRET });
        return h.digest({ outputEncoding: crypto.Encoding.HEX });
      }
      function verify(empid, ts, sig) {
        if (!empid || !ts || !sig) return false;
        if (Math.abs(Date.now() - parseInt(ts, 10)) > TOKEN_TTL_MS) return false;
        try { return sign(empid, ts) === sig; } catch (e) { log.error('verify token', e); return false; }
      }

      if (context.request.method === 'GET') {

        var q = context.request.parameters || {};
        var typeParam = (q.type || 3).toString();
        var showTopFilters = (typeParam !== '4');
        var formName = 'MI Reorder Tool (' + (typeParam == 4 ? 'Basic' : 'Admin') + ')';
        var mode = (q.mode || '').toString().toLowerCase();
        var cronTs = q.cronts || '';
        var cronSig = q.cronsig || '';
        var isCron = (mode === 'cron' && verifyCron(cronTs, cronSig));

        var form = ui.createForm({ title: formName });

        // --- Signed params (preferred)
        var empid = q.empid || '';
        var ts = q.ts || '';
        var sig = q.sig || '';

        // --- Legacy fallback
        var selectedEmp = q.custpage_id || '';

        if (empid && ts && sig && verify(empid, ts, sig)) {
          selectedEmp = empid;
        }
        if (!selectedEmp && !isCron) {
          context.response.write(`
            <html><head>
            <script>setTimeout(function(){ window.location.href = ${JSON.stringify(portalUrl)}; }, 1200);</script>
            <style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>
            </head><body><div class="message">Login Required</div></body></html>
          `);
          return;
        }

        var empField = form.addField({ id: 'custpage_empid', type: ui.FieldType.SELECT, label: 'Current Employee', source: 'employee' });
        empField.defaultValue = selectedEmp;
        empField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

        var tsField = form.addField({ id: 'custpage_ts', type: ui.FieldType.TEXT, label: 'ts' });
        tsField.defaultValue = ts || '';
        tsField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

        var sigField = form.addField({ id: 'custpage_sig', type: ui.FieldType.TEXT, label: 'sig' });
        sigField.defaultValue = sig || '';
        sigField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

        var fileField = form.addField({ id: 'custpage_file_id', type: ui.FieldType.TEXT, label: 'File ID' });
        fileField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

        var htmlField = form.addField({ id: 'custpage_excel_html', type: ui.FieldType.INLINEHTML, label: 'Excel Table' });

        var selectedField = form.addField({ id: 'custpage_selected', type: ui.FieldType.LONGTEXT, label: 'Selected Rows JSON' });
        selectedField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

        form.addSubmitButton({ label: 'Submit' });

        // Get latest file from folder
        var fileId = null;
        var fileId2 = null;
        var fileId3 = null;
        var folderSearchObj = search.create({
          type: "folder",
          filters: [
            ["internalid", "anyof", "402335", "402334", "413248"],
            "AND",
            ["file.documentsize", "greaterthan", "5"]
          ],
          columns: [
            search.createColumn({ name: "internalid", summary: "GROUP" }),
            search.createColumn({ name: "internalid", join: "file", summary: "MAX" })
          ]
        });

        folderSearchObj.run().each(function (result) {
          log.debug('result', result);
          var folderID = result.getValue({ name: "internalid", summary: "GROUP" });
          if (folderID == 402335)
            fileId = result.getValue({ name: "internalid", join: "file", summary: "MAX" });
          if (folderID == 402334)
            fileId2 = result.getValue({ name: "internalid", join: "file", summary: "MAX" });
          if (folderID == 413248)
            fileId3 = result.getValue({ name: "internalid", join: "file", summary: "MAX" });
          return true;
        });

        if (!fileId) {
          htmlField.defaultValue = '<p style="color:red;">No file found in the specified folder.</p>';
          context.response.writePage(form);
          return;
        }

        var fileObj = file.load({ id: fileId });
        var content = fileObj.getContents();
        var rows = content.split('\n');
        log.debug('rows', rows.length);

        var fileObj2 = file.load({ id: fileId2 });
        var content2 = fileObj2.getContents();
        var rows2 = content2.split('\n');

        var fileObj3 = file.load({ id: fileId3 });
        var content3 = fileObj3.getContents();
        var rows3 = content3.split('\n');

        rows3.shift();
        rows2.shift();
        rows = rows.concat(rows2);
        rows = rows.concat(rows3);
        log.debug('rows', rows.length);

        var newContent = [];
        const ITEM_INDEX = 2;
        const VENDOR_INDEX = 36;
        const BRAND_CATEGORY_INDEX = 34;
        const BRAND_INDEX = 33;
        const DEPT_INDEX = 53;
        const POL_INDEX = 37;
        const PRODUCT_INDEX = 56;
        const NEW_COL_SHIFT = 2; // 2 new columns inserted after index 5

        const uniqueItems = new Set();
        const uniqueVendors = new Set();
        const uniqueBrands = new Set();
        const uniqueBrandCategories = new Set();
        const uniqueDepartments = new Set();
        const uniquePOL = new Set();
        const uniquePRODUCT = new Set();

        // ============================================================
        //  HTML + CSS (Optimized UI)
        // ============================================================
        let html = `
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');

  /* ── Reset & base ── */
  *, *::before, *::after { box-sizing: border-box; }

  .rr-root {
    font-family: 'DM Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    font-size: 13px;
    color: #1e293b;
    line-height: 1.4;
    display: flex;
    flex-direction: column;
    height: calc(100vh - var(--rr-top-offset, 80px));
    min-height: 400px;
    overflow: hidden;
  }

  /* ── Toolbar ── */
  .rr-toolbar {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 10px;
    padding: 12px 16px;
    margin-bottom: 0;
    background: linear-gradient(135deg, #f8fafc 0%, #eef2f7 100%);
    border: 1px solid #e2e8f0;
    border-bottom: none;
    border-radius: 10px 10px 0 0;
    flex-shrink: 0;
  }

  .rr-toolbar-right {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 10px;
  }

  /* ── Pill badges ── */
  .rr-pill {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 5px 14px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.3px;
    border: 1px solid;
    transition: transform 0.15s ease;
  }
  .rr-pill:hover { transform: translateY(-1px); }

  .rr-pill--cubic {
    background: #eff6ff;
    border-color: #bfdbfe;
    color: #1e40af;
  }
  .rr-pill--weight {
    background: #f0fdf4;
    border-color: #bbf7d0;
    color: #166534;
  }
  .rr-pill__val {
    font-variant-numeric: tabular-nums;
    font-weight: 700;
    font-size: 13px;
  }

  /* ── Download button ── */
  .rr-dl-btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 6px 14px;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    background: #fff;
    color: #334155;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.15s ease;
  }
  .rr-dl-btn:hover { background: #f1f5f9; border-color: #94a3b8; }
  .rr-dl-btn img { width: 15px; height: 15px; }

  /* ── Filter panel (top-level dropdown) ── */
  .rr-filter-wrap {
    position: relative;
    display: inline-block;
  }
  .rr-filter-trigger {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 6px 14px;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    background: #fff;
    color: #334155;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.15s ease;
    user-select: none;
  }
  .rr-filter-trigger:hover { background: #f1f5f9; border-color: #94a3b8; }
  .rr-filter-trigger svg { width: 14px; height: 14px; }

  .rr-filter-dropdown {
    display: none;
    position: absolute;
    top: calc(100% + 6px);
    left: 0;
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 16px;
    width: 680px;
    box-shadow: 0 12px 40px rgba(0,0,0,0.1);
    z-index: 6000;
  }
  .rr-filter-dropdown.open { display: block; }

  .rr-filter-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 12px;
  }
  .rr-filter-grid label {
    display: block;
    font-weight: 600;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: #64748b;
    margin-bottom: 4px;
  }
  .rr-filter-grid select {
    width: 100%;
    padding: 6px 8px;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    font-size: 12px;
    font-family: inherit;
    background: #f8fafc;
    transition: border-color 0.15s;
  }
  .rr-filter-grid select:focus {
    outline: none;
    border-color: #60a5fa;
    box-shadow: 0 0 0 3px rgba(96,165,250,0.15);
  }

  /* ── Scrollbar sync ── */
  .rr-hscroll {
    height: 14px;
    overflow-x: auto;
    overflow-y: hidden;
    border: 1px solid #e2e8f0;
    border-bottom: 0;
    border-radius: 0;
    flex-shrink: 0;
  }
  .rr-hscroll-inner { height: 1px; }

  /* ── Table viewport ── */
  .rr-viewport {
    max-width: calc(100vw - 10px);
    width: 100%;
    flex: 1 1 0%;
    display: flex;
    flex-direction: column;
    min-height: 0;           /* critical for flex child to shrink */
    overflow: hidden;
  }
  .rr-table-wrap {
    flex: 1 1 0%;
    min-height: 0;           /* allows shrinking inside flex */
    overflow: auto;
    border: 1px solid #e2e8f0;
    border-radius: 0 0 10px 10px;
  }

  /* ── Table ── */
  #excelTable {
    border-collapse: separate;
    border-spacing: 0;
    width: max-content;
    min-width: 100%;
    table-layout: auto;
  }

  #excelTable th,
  #excelTable td {
    border-bottom: 1px solid #f1f5f9;
    border-right: 1px solid #f1f5f9;
    padding: 7px 12px;
    white-space: nowrap;
    font-size: 12px;
    background: #fff;
    vertical-align: middle;
  }

  /* Header */
  #excelTable thead th {
    position: sticky;
    top: 0;
    background: #f8fafc;
    color: #475569;
    font-weight: 600;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.4px;
    border-bottom: 2px solid #e2e8f0;
    z-index: 9;
    padding: 9px 12px;
  }

  /* Alternating rows */
  #excelTable tbody tr:nth-child(even) td { background: #fafbfc; }
  #excelTable tbody tr:hover td { background: #f0f7ff !important; }

  /* Sticky columns */
  .sticky-col {
    position: sticky;
    background: inherit;
    background-clip: padding-box;
    z-index: 11;
  }
  thead th.sticky-col { z-index: 30; }
  .sticky-col.sep-left {
    border-left-color: transparent;
    box-shadow: inset 1px 0 0 #e2e8f0;
  }

  /* ── Inputs inside table ── */
  #excelTable input[type="number"],
  #excelTable input[type="text"] {
    width: 100%;
    box-sizing: border-box;
    padding: 4px 8px;
    border: 1px solid #e2e8f0;
    border-radius: 5px;
    font-family: inherit;
    font-size: 12px;
    transition: border-color 0.15s, box-shadow 0.15s;
    background: #fff;
  }
  #excelTable input:focus {
    outline: none;
    border-color: #60a5fa;
    box-shadow: 0 0 0 2px rgba(96,165,250,0.15);
  }
  #excelTable input[type="checkbox"] {
    width: 16px; height: 16px;
    cursor: pointer;
    accent-color: #3b82f6;
  }

  /* ── Column filter button in header ── */
  .th-filter-wrap {
    display: inline-flex;
    align-items: center;
    gap: 5px;
  }
  .th-filter-btn {
    cursor: pointer;
    border: 1px solid transparent;
    background: transparent;
    padding: 1px 4px;
    border-radius: 4px;
    line-height: 1;
    font-size: 10px;
    color: #94a3b8;
    user-select: none;
    transition: all 0.15s;
  }
  .th-filter-btn:hover { background: #e2e8f0; color: #475569; }
  .th-filter-active .th-filter-btn {
    background: #dbeafe;
    color: #2563eb;
    border-color: #93c5fd;
  }

  /* ── Column filter panel (portal) ── */
  .th-filter-panel {
    position: fixed;
    top: 0; left: 0;
    width: 270px;
    max-height: 360px;
    overflow: auto;
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 12px;
    box-shadow: 0 12px 40px rgba(0,0,0,0.12);
    z-index: 9999;
    font-family: 'DM Sans', sans-serif;
  }
  .th-filter-panel .hdr {
    font-weight: 700;
    font-size: 12px;
    color: #1e293b;
    margin-bottom: 8px;
  }
  .th-filter-panel input[type="search"] {
    width: 100%;
    padding: 7px 10px;
    box-sizing: border-box;
    margin-bottom: 8px;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    font-size: 12px;
    font-family: inherit;
    transition: border-color 0.15s;
  }
  .th-filter-panel input[type="search"]:focus {
    outline: none;
    border-color: #60a5fa;
    box-shadow: 0 0 0 2px rgba(96,165,250,0.12);
  }
  .th-filter-panel .list {
    max-height: 200px;
    overflow: auto;
    border: 1px solid #f1f5f9;
    border-radius: 6px;
    padding: 6px;
  }
  .th-filter-panel .row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 3px 4px;
    font-size: 12px;
    border-radius: 4px;
    transition: background 0.1s;
  }
  .th-filter-panel .row:hover { background: #f8fafc; }
  .th-filter-panel .row input[type="checkbox"] {
    width: 14px; height: 14px;
    accent-color: #3b82f6;
  }
  .th-filter-panel .actions {
    display: flex;
    justify-content: space-between;
    gap: 6px;
    margin-top: 10px;
  }
  .th-filter-panel .actions button {
    padding: 4px 10px;
    font-size: 11px;
    font-weight: 600;
    border-radius: 6px;
    border: 1px solid #e2e8f0;
    background: #fff;
    cursor: pointer;
    font-family: inherit;
    transition: all 0.15s;
  }
  .th-filter-panel .actions button:hover { background: #f1f5f9; }
  .th-filter-panel .actions button[data-act="apply"] {
    background: #3b82f6;
    color: #fff;
    border-color: #3b82f6;
  }
  .th-filter-panel .actions button[data-act="apply"]:hover { background: #2563eb; }
  .th-filter-panel .actions button[data-act="clear"] {
    color: #dc2626;
    border-color: #fecaca;
  }
  .th-filter-panel .actions button[data-act="clear"]:hover { background: #fef2f2; }

  /* ── Calculated cell highlight ── */
  .rr-calc-cell {
    background: #eff6ff !important;
    font-weight: 600;
    color: #1e40af;
  }
</style>

<div class="rr-root">

<!-- ── Toolbar ── -->
<div class="rr-toolbar">
  <button id="downloadCsvBtn" class="rr-dl-btn" type="button">
    <img src="https://cdn-icons-png.flaticon.com/512/10630/10630240.png" alt="csv" />
    Export CSV
  </button>
  ${showTopFilters ? `
  <div class="rr-filter-wrap" id="filterWrapper">
    <div class="rr-filter-trigger" id="filterToggle">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 3H2l8 9.46V19l4 2v-8.54L22 3z"/></svg>
      Filters
    </div>
    <div class="rr-filter-dropdown" id="filterPanel">
      <div class="rr-filter-grid">
        <div><label for="itemFilter">Item</label><select id="itemFilter" multiple size="5"></select></div>
        <div><label for="vendorFilter">Vendor</label><select id="vendorFilter" multiple size="5"></select></div>
        <div><label for="brandFilter">Brand</label><select id="brandFilter" multiple size="5"></select></div>
        <div><label for="brandCatFilter">Brand Category</label><select id="brandCatFilter" multiple size="5"></select></div>
        <div><label for="deptFilter">Department</label><select id="deptFilter" multiple size="5"></select></div>
        <div><label for="polFilter">P.O.L</label><select id="polFilter" multiple size="5"></select></div>
        <div><label for="productFilter">Product Type</label><select id="productFilter" multiple size="5"></select></div>
      </div>
    </div>
  </div>
  ` : ''}

  <div class="rr-toolbar-right">
    <div class="rr-pill rr-pill--cubic" title="Sum of Item Cubic Space for all selected rows">
      <span>Total Cubic Space</span>
      <span class="rr-pill__val" id="totalCubicValue">0</span>
    </div>
    <div class="rr-pill rr-pill--weight" title="Sum of Item Weight for all selected rows">
      <span>Total Weight</span>
      <span class="rr-pill__val" id="totalWeightValue">0</span>
    </div>
  </div>
</div>

<!-- ── Table area ── -->
<div class="rr-viewport">
  <div class="rr-hscroll" id="topScroll">
    <div class="rr-hscroll-inner" id="topScrollInner"></div>
  </div>
  <div class="rr-table-wrap" id="tableContainer">
    <table id="excelTable">
      <thead><tr>
`;

        const headers = rows[0].split(/,(?=(?:[^\"]*"[^"]*")*[^\"]*$)/);

        function norm(h) {
          return String(h || '')
            .replace(/^[\uFEFF\s"]+|[\s"]+$/g, '')
            .replace(/\s+/g, ' ')
            .toLowerCase();
        }

        const rawHeaders = headers;
        const normalized = rawHeaders.map(norm);

        let adminCsvIndex = normalized.indexOf('admin portal');
        if (adminCsvIndex < 0) {
          adminCsvIndex = normalized.findIndex(h =>
            h === 'admin' || h.startsWith('admin portal') || h.includes('admin portal')
          );
        }

        const truncateAfterAdmin = (typeParam === '4');
        const removeJustAdmin = (typeParam === '3');

        const headersForOutput = (() => {
          if (adminCsvIndex >= 0) {
            if (truncateAfterAdmin) return rawHeaders.slice(0, adminCsvIndex);
            if (removeJustAdmin) return rawHeaders.filter((_, i) => i !== adminCsvIndex);
          }
          return rawHeaders.slice();
        })();
        // Insert 2 new PO qty columns after index 5
        headersForOutput.splice(6, 0, '"Qty Ordered Last Year"', '"Qty Ordered This Year"');
        headersForOutput.push('Filter');
        newContent.push(headersForOutput);

        // Build header cells
        // 5 control columns: Select, Order Qty, Month of Stock, Item Space, Weight
        html += '<th class="sticky-col" style="left:0;">Select</th>'
          + '<th class="sticky-col">Order Qty</th>'
          + '<th class="sticky-col">Month of Stock</th>'
          + '<th class="sticky-col">Item Space</th>'
          + '<th class="sticky-col">Weight</th>';

        var dataColStart = 5;
        var visibleIdx = 0;

        headersForOutput.forEach(function (value, outIdx) {
          var domIndex = dataColStart + visibleIdx;
          visibleIdx++;

          value = String(value || '').replace(/"/g, '').trim();
          var safeText = value || ('Col ' + (outIdx + 1));

          html += '<th data-domidx="' + domIndex + '">'
            + '<span class="th-filter-wrap">'
            + '<span class="th-label">' + safeText + '</span>'
            + '<button class="th-filter-btn" type="button" data-col="' + domIndex + '">&#9662;</button>'
            + '</span>'
            + '</th>';
        });

        html += '</tr></thead><tbody>';

        let redRows = [];
        let blackRows = [];

        function csvIndexIsExposed(idx) {
          if (idx == adminCsvIndex || idx > adminCsvIndex) idx = idx + 1;
          return idx;
        }

        var balances = getInventoryBalanceMap();
        log.debug('Inventory Balance Map', balances);
        var poQtyMap = getPurchaseOrderQtyMap();
        log.debug('PO Qty Map', poQtyMap);

        var commMap = getCommBackMap();
        log.debug('Commited Qty Map', commMap);
        var alreadyExsist = {};

        rows.slice(1).forEach(function (row) {
          if (!row || row.trim() === '') return;

          const columns = row.split(/,(?=(?:[^\"]*"[^"]*")*[^\"]*$)/);
          if (!columns.length) return;

          if (typeParam === '4' && adminCsvIndex >= 0) {
            var adminCellVal = (columns[adminCsvIndex] || '').replace(/"/g, '').trim().toLowerCase();
            if (adminCellVal !== 'admin portal') return;
          }

          let calcCols = columns.slice();
          const monthAvg = parseFloat(calcCols[csvIndexIsExposed(11)]);
          const itemid = calcCols[csvIndexIsExposed(3)];

          if (alreadyExsist.hasOwnProperty(itemid) || !itemid) return;
          alreadyExsist[itemid] = true;

          if (columns[csvIndexIsExposed(ITEM_INDEX)]) uniqueItems.add(columns[csvIndexIsExposed(ITEM_INDEX)].replace(/"/g, '').trim());
          if (columns[csvIndexIsExposed(VENDOR_INDEX)]) uniqueVendors.add(columns[csvIndexIsExposed(VENDOR_INDEX)].replace(/"/g, '').trim());
          if (columns[csvIndexIsExposed(BRAND_INDEX)]) uniqueBrands.add(columns[csvIndexIsExposed(BRAND_INDEX)].replace(/"/g, '').trim());
          if (columns[csvIndexIsExposed(BRAND_CATEGORY_INDEX)]) uniqueBrandCategories.add(columns[csvIndexIsExposed(BRAND_CATEGORY_INDEX)].replace(/"/g, '').trim());
          if (columns[csvIndexIsExposed(DEPT_INDEX)]) uniqueDepartments.add(columns[csvIndexIsExposed(DEPT_INDEX)].replace(/"/g, '').trim());
          if (columns[csvIndexIsExposed(POL_INDEX)]) uniquePOL.add(columns[csvIndexIsExposed(POL_INDEX)].replace(/"/g, '').trim());
          if (columns[csvIndexIsExposed(PRODUCT_INDEX)]) uniquePRODUCT.add(columns[csvIndexIsExposed(PRODUCT_INDEX)].replace(/"/g, '').trim());

          var val43 = parseFloat(calcCols[csvIndexIsExposed(26)] || 0);
          var val41 = parseFloat(calcCols[csvIndexIsExposed(21)] || 0);
          var diff = val43 - val41;
          diff = Math.abs(diff);
          calcCols[csvIndexIsExposed(26)] = diff === 0 ? "" : '"' + diff + '"';

          let good     = 0;
          let bad      = 0;
          let hold     = 0;
          let inspect  = 0;
          let label    = 0;
          let total    = 0;
          let avail    = 0;
          let onH      = 0;
          let committedQty = 0;
          let backOrdered = 0;
          let col14Val = calcCols[csvIndexIsExposed(19)];
          let col12Val = parseFloat(calcCols[csvIndexIsExposed(12)] || 0);
          let col19Val = diff;
          
          if (commMap[itemid]) {
            committedQty    = parseFloat(commMap[itemid].qtyComm);
            backOrdered     = parseFloat(commMap[itemid].qtyBack);
          }

          if (balances[itemid]) {
            good    = parseFloat(balances[itemid].good);
            bad     = balances[itemid].bad;
            hold    = balances[itemid].hold;
            inspect = balances[itemid].inspect;
            label   = balances[itemid].label;
            onH     = parseFloat(balances[itemid].onH);
            total   = parseFloat(balances[itemid].total);
            avail   = parseFloat(balances[itemid].avail);
          }

          if (itemid == 1163) {
            log.debug('committedQty', committedQty);
            log.debug('onH', onH);
            log.debug('good', good);
            log.debug('avail', avail);
            log.debug('total', total);
            log.debug('col12Val', col12Val);
            log.debug('balances', balances[itemid]);
            log.debug('Test', good - col12Val);
          }

          

          var availtoProm = good - committedQty;

          calcCols[calcCols.length] = 'Black';
          calcCols[csvIndexIsExposed(12)] = '"' + availtoProm + '"';
          calcCols[csvIndexIsExposed(14)] = '"' + good + '"';
          calcCols[csvIndexIsExposed(15)] = '"' + bad + '"';
          calcCols[csvIndexIsExposed(16)] = '"' + inspect + '"';
          calcCols[csvIndexIsExposed(17)] = '"' + label + '"';
          calcCols[csvIndexIsExposed(18)] = '"' + hold + '"';
          calcCols[csvIndexIsExposed(19)] = '"' + total + '"';
          calcCols[csvIndexIsExposed(20)] = '"' + ((parseFloat(total)) / monthAvg).toFixed(2) + '"';
          calcCols[csvIndexIsExposed(25)] = '"' + (((parseFloat(total)) + parseFloat(val41 || 0)) / monthAvg).toFixed(2) + '"';
          calcCols[csvIndexIsExposed(24)] = '"' + ((parseFloat(total)) + parseFloat(val41 || 0)).toFixed(2) + '"';
          calcCols[csvIndexIsExposed(28)] = '"' + ((parseFloat(total) + parseFloat(val43 || 0)) / monthAvg).toFixed(2) + '"';
          calcCols[csvIndexIsExposed(27)] = '"' + (parseFloat(total) + parseFloat(val43 || 0)).toFixed(2) + '"';
          calcCols[csvIndexIsExposed(32)] = '"' + avail + '"';

          function normalizeMovement(val) {
            if (val === null || val === undefined) return "No Movement";
            var s = String(val).replace(/,/g, "").trim();
            if (s === "") return "No Movement";
            var n = parseFloat(s);
            if (!isFinite(n)) return "No Movement";
            if (n <= 0) return "No Movement";
            return n.toFixed(2);
          }

          var col9 = normalizeMovement(monthAvg);
          const qtytotal = diff + safeParseFloat(calcCols[csvIndexIsExposed(21)]) + parseFloat(avail) - parseFloat(col12Val);
          const stockingQty = Math.ceil(parseFloat(calcCols[csvIndexIsExposed(11)]) * 4.5);
          calcCols[csvIndexIsExposed(11)] = '"' + col9 + '"';

          calcCols[csvIndexIsExposed(63)] = committedQty;
          calcCols[csvIndexIsExposed(64)] = backOrdered;

          calcCols[csvIndexIsExposed(66)] = calcCols[csvIndexIsExposed(11)];
          calcCols[csvIndexIsExposed(67)] = stockingQty;

          let recommendedQty = 0;
          if (qtytotal < stockingQty) {
            recommendedQty = (stockingQty - qtytotal).toFixed(2);
          }

          const monthsStock = (diff + parseFloat(avail) + safeParseFloat(calcCols[csvIndexIsExposed(21)])) / monthAvg;

          if (calcCols[1] == 0) calcCols[1] = '""';
          calcCols[1] = '"' + recommendedQty + '"';

          let rowStyle = '';
          if (!isNaN(monthsStock) && monthsStock <= 4.5) {
            rowStyle = 'color: #dc2626; font-weight: 500;';
            calcCols[calcCols.length - 1] = 'Red';
          }

          var statusVal = (calcCols.length ? calcCols[calcCols.length - 1] : '') || '';
          var baseCols = calcCols.slice(0, -1);

          // Insert "Qty - Last Year" and "Qty - This Year" at position 6 & 7
          // Must match headersForOutput.splice(6, ...) exactly
          var poData = poQtyMap[itemid] || { lastYear: 0, thisYear: 0 };
          baseCols.splice(6, 0,
            '"' + poData.lastYear + '"',
            '"' + poData.thisYear + '"'
          );

          let displayCols;
          // adminCsvIndex is based on original array; baseCols now has +2 cols at position 6
          var adjustedAdminIdx = (adminCsvIndex >= 0 && adminCsvIndex >= 6) ? adminCsvIndex + 2 : adminCsvIndex;
          if (adjustedAdminIdx >= 0 && truncateAfterAdmin) {
            displayCols = baseCols.slice(0, adjustedAdminIdx);
          } else if (adjustedAdminIdx >= 0 && removeJustAdmin) {
            displayCols = baseCols.filter((_, i) => i !== adjustedAdminIdx);
          } else {
            displayCols = baseCols.slice();
          }
          displayCols.push(statusVal);

          function cleanCsvValue(value) {
            let v = value == null ? '' : String(value).trim();
            if (v.length >= 2 && v[0] === '"' && v[v.length - 1] === '"') {
              const inner = v.slice(1, -1);
              if (/[",\r\n]/.test(inner) || /""/.test(inner)) {
                return '"' + inner.replace(/"/g, '""') + '"';
              }
              v = inner;
            }
            if ((v.startsWith('"') && !v.endsWith('"')) || (!v.startsWith('"') && v.endsWith('"'))) {
              v = v.replace(/^"+|"+$/g, '');
            }
            if ((v.startsWith('"') && !v.endsWith('"')) || (!v.startsWith('"') && v.endsWith('"'))) {
              v = v.replace(/^"+|"+$/g, '');
            }
            if (/[",\r\n]/.test(v)) {
              v = '"' + v.replace(/"/g, '""') + '"';
            }
            if (v === '- None -' || v === 'NaN' || v === 'Infinity') v = '';
            if (v === '.00') v = '0.00';
            return v;
          }

          const cleanedCols = [];
          displayCols.forEach(function (value) {
            cleanedCols.push(cleanCsvValue(value));
          });
          newContent.push(cleanedCols);

          // --- Render row ---
          let rowId = newContent.length - 1;
          let rowHtml = '<tr style="' + rowStyle + '">';
          rowHtml += '<td class="sticky-col" style="left:0;text-align:center;"><input type="checkbox" name="row_select_' + rowId + '" /></td>';
          rowHtml += '<td class="sticky-col"><input type="number" name="qty_input_' + rowId + '" min="0" value="' + recommendedQty + '" style="width:80px;" /></td>';
          rowHtml += '<td class="sticky-col"></td>';  // Month of Stock (computed by JS)
          rowHtml += '<td class="sticky-col"></td>';  // Item Space (computed by JS)
          rowHtml += '<td class="sticky-col"></td>';  // Weight (computed by JS)

          var first = true;
          displayCols.forEach(function (value) {
            if (first) {
              value = String(value || '').replace(/"/g, '').trim();
              if (value === '- None -' || value === 'NaN' || value === 'Infinity') value = '';
              rowHtml += '<td><input type="text" name="memo_input_' + rowId + '" value="' + value + '" /></td>';
            } else {
              value = String(value || '').replace(/"/g, '').trim();
              if (value === '- None -' || value === 'NaN' || value === 'Infinity') value = '';
              if (value === '.00') value = '0.00';
              rowHtml += '<td>' + value + '</td>';
            }
            first = false;
          });

          rowHtml += '</tr>';

          var itemIdNum = parseInt(itemid, 10) || 0;

          if (rowStyle.includes('dc2626')) {
            redRows.push({ itemId: itemIdNum, html: rowHtml });
          } else {
            blackRows.push({ itemId: itemIdNum, html: rowHtml });
          }
        });

        redRows.sort(function (a, b) { return a.itemId - b.itemId; });
        blackRows.sort(function (a, b) { return a.itemId - b.itemId; });

        html += redRows.map(function (r) { return r.html; }).join('')
          + blackRows.map(function (r) { return r.html; }).join('');

        html += '</tbody></table></div></div></div>';

        // Filter portal
        html += '<div id="filter-portal"></div>';

        // ============================================================
        //  INLINE SCRIPTS (replaces client script + fixes filters)
        // ============================================================
        html += `
<script>
(function(){
  "use strict";

  // ==== Constants for cell indices (TD positions in each row) ====
  var MONTH_CELL_INDEX = 2;
  var CUBIC_CELL_INDEX = 3;
  var WEIGHT_CELL_INDEX = 4;

  // Source columns in the table (data column TD indices)
  // +2 shift for all >= old 6 due to Qty Last Year / This Year inserted after col 5
  var SRC_PER_CUBIC_IDX = 38; 
  var SRC_PER_WGT_IDX = 36;   
  var SRC_PER_WGT_UNIT_IDX = 37;
  var SRC_PER_AVAIL_IDX = 39;   
  var SRC_MONTH_AVG_IDX_1 = 18; 

  // ── Helpers ──
  function toNum(v) {
    var n = parseFloat((v || '').toString().replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }

  function perUnitWeightToKg(value, unit) {
    if (!value || isNaN(value)) return 0;
    if (unit === 'g' || unit === 'gram' || unit === 'grams') return value / 1000;
    if (unit === 'kg' || unit === 'kilogram' || unit === 'kilograms') return value;
    return value * 0.453592;
  }

  function formatNumber(n, maxFrac) {
    var num = typeof n === 'number' ? n : toNum(n);
    return num.toLocaleString(undefined, { maximumFractionDigits: maxFrac || 2 });
  }

  function styleCalcCell(el) {
    el.className = (el.className || '').replace(/\\brr-calc-cell\\b/g, '') + ' rr-calc-cell';
  }
  function clearCalcCell(el) {
    el.className = (el.className || '').replace(/\\brr-calc-cell\\b/g, '').trim();
    el.textContent = '';
  }

  // ── Per-row calculation ──
  function updateRowCells(row, qtyInput) {
    var qtyOrdered = toNum(qtyInput.value);
    var cells = row.querySelectorAll('td');
    if (cells.length <= WEIGHT_CELL_INDEX) return;

    var monthQty = parseFloat(cells[SRC_MONTH_AVG_IDX_1] ? cells[SRC_MONTH_AVG_IDX_1].textContent : 0) || 0;
    var inTransit = parseFloat(cells[28] ? cells[28].textContent : 0) || 0;
    var onOrder = parseFloat(cells[33] ? cells[33].textContent : 0) || 0;
    var avail = parseFloat(cells[SRC_PER_AVAIL_IDX] ? cells[SRC_PER_AVAIL_IDX].textContent : 0) || 0;

    console.log('Values', {monthQty, inTransit, onOrder, avail})

    var totalForMos = inTransit + onOrder + avail + qtyOrdered;
    var monthCell = cells[MONTH_CELL_INDEX];
    monthCell.textContent = monthQty ? (totalForMos / monthQty).toFixed(2) : '0.00';
    styleCalcCell(monthCell);

    var perCubic = toNum(cells[SRC_PER_CUBIC_IDX] ? cells[SRC_PER_CUBIC_IDX].textContent : 0);
    var rowCubic = qtyOrdered * (isNaN(perCubic) ? 0 : perCubic);
    var cubicCell = cells[CUBIC_CELL_INDEX];
    cubicCell.textContent = formatNumber(rowCubic, 2);
    styleCalcCell(cubicCell);

    var perWgtVal = toNum(cells[SRC_PER_WGT_IDX] ? cells[SRC_PER_WGT_IDX].textContent : 0);
    var perWgtUnit = (cells[SRC_PER_WGT_UNIT_IDX] ? cells[SRC_PER_WGT_UNIT_IDX].textContent : '').toLowerCase().trim();
    var perWgtKg = perUnitWeightToKg(perWgtVal, perWgtUnit);
    var rowWgtKg = qtyOrdered * perWgtKg;
    var weightCell = cells[WEIGHT_CELL_INDEX];
    weightCell.textContent = formatNumber(rowWgtKg, 2) + ' kg';
    styleCalcCell(weightCell);
  }

  function clearRowCells(row) {
    var cells = row.querySelectorAll('td');
    [MONTH_CELL_INDEX, CUBIC_CELL_INDEX, WEIGHT_CELL_INDEX].forEach(function(idx) {
      if (cells[idx]) clearCalcCell(cells[idx]);
    });
  }

  // ── Totals across all checked rows ──
  function updateTotals() {
    var totalCubic = 0;
    var totalWeightKg = 0;

    document.querySelectorAll('input[type="checkbox"][name^="row_select_"]').forEach(function(cb) {
      if (!cb.checked) return;
      var row = cb.closest('tr');
      var rowId = cb.name.split('_')[2];
      var qtyInput = document.querySelector('input[name="qty_input_' + rowId + '"]');
      if (!row || !qtyInput) return;

      var cells = row.querySelectorAll('td');
      var qty = toNum(qtyInput.value);

      var perCubic = toNum(cells[SRC_PER_CUBIC_IDX] ? cells[SRC_PER_CUBIC_IDX].textContent : 0);
      totalCubic += qty * (isNaN(perCubic) ? 0 : perCubic);

      var perWgtVal = toNum(cells[SRC_PER_WGT_IDX] ? cells[SRC_PER_WGT_IDX].textContent : 0);
      var perWgtUnit = (cells[SRC_PER_WGT_UNIT_IDX] ? cells[SRC_PER_WGT_UNIT_IDX].textContent : '').toLowerCase().trim();
      totalWeightKg += qty * perUnitWeightToKg(perWgtVal, perWgtUnit);
    });

    var cubicEl = document.getElementById('totalCubicValue');
    if (cubicEl) cubicEl.textContent = formatNumber(totalCubic, 2);
    var wgtEl = document.getElementById('totalWeightValue');
    if (wgtEl) wgtEl.textContent = formatNumber(totalWeightKg, 2);
  }

  // ── Bind checkbox & qty events (replaces client script pageInit) ──
  function initRowEvents() {
    var checkboxes = document.querySelectorAll('input[type="checkbox"][name^="row_select_"]');
    checkboxes.forEach(function(checkbox) {
      var rowId = checkbox.name.split('_')[2];
      var qtyInput = document.querySelector('input[name="qty_input_' + rowId + '"]');
      var row = checkbox.closest('tr');
      if (!qtyInput || !row) return;

      checkbox.addEventListener('change', function() {
        if (checkbox.checked) {
          updateRowCells(row, qtyInput);
        } else {
          clearRowCells(row);
        }
        updateTotals();
      });

      qtyInput.addEventListener('input', function() {
        if (checkbox.checked) {
          updateRowCells(row, qtyInput);
          updateTotals();
        }
      });
    });
    updateTotals();
  }

  // ═════════════════════════════════════════════════════════════
  //  TOP-LEVEL FILTER DROPDOWN
  // ═════════════════════════════════════════════════════════════
  function initTopFilters() {
    var wrapper = document.getElementById('filterWrapper');
    var panel = document.getElementById('filterPanel');
    if (!wrapper || !panel) return;

    var timer;
    wrapper.addEventListener('mouseenter', function() {
      clearTimeout(timer);
      panel.classList.add('open');
    });
    wrapper.addEventListener('mouseleave', function() {
      timer = setTimeout(function() { panel.classList.remove('open'); }, 200);
    });

    var itemSet = ${JSON.stringify([...uniqueItems])};
    var vendorSet = ${JSON.stringify([...uniqueVendors])};
    var brandSet = ${JSON.stringify([...uniqueBrands])};
    var brandCatSet = ${JSON.stringify([...uniqueBrandCategories])};
    var deptSet = ${JSON.stringify([...uniqueDepartments])};
    var polSet = ${JSON.stringify([...uniquePOL])};
    var productSet = ${JSON.stringify([...uniquePRODUCT])};

    var itemSelect = document.getElementById('itemFilter');
    var vendorSelect = document.getElementById('vendorFilter');
    var brandSelect = document.getElementById('brandFilter');
    var brandCatSelect = document.getElementById('brandCatFilter');
    var deptSelect = document.getElementById('deptFilter');
    var polSelect = document.getElementById('polFilter');
    var productSelect = document.getElementById('productFilter');

    function populateSelect(selectEl, arr) {
      if (!selectEl) return;
      arr.forEach(function(val) {
        var opt = document.createElement('option');
        opt.value = val;
        opt.textContent = val;
        selectEl.appendChild(opt);
      });
    }
    populateSelect(itemSelect, itemSet);
    populateSelect(vendorSelect, vendorSet);
    populateSelect(brandSelect, brandSet);
    populateSelect(brandCatSelect, brandCatSet);
    populateSelect(deptSelect, deptSet);
    populateSelect(polSelect, polSet);
    populateSelect(productSelect, productSet);

    function getSelectedValues(selectEl) {
      if (!selectEl) return [];
      var selected = [];
      for (var i = 0; i < selectEl.options.length; i++) {
        if (selectEl.options[i].selected) selected.push(selectEl.options[i].value.toLowerCase());
      }
      return selected;
    }

    // Column mapping: CSV index -> DOM td index
    // The 5 control cols shift everything by 5
    function vizIdxFromCsv(csvIdx) {
      var controlOffset = 5;
      var hasAdmin = (typeof window.__ADMIN_IDX__ === 'number' && window.__ADMIN_IDX__ >= 0);
      if (window.__TRUNC_AFTER__ && hasAdmin) {
        if (csvIdx >= window.__ADMIN_IDX__) return -1;
        return controlOffset + csvIdx;
      }
      if (window.__ADMIN_REMOVED__ && hasAdmin) {
        var shift = (csvIdx >= window.__ADMIN_IDX__) ? 1 : 0;
        return controlOffset + csvIdx - shift;
      }
      return controlOffset + csvIdx;
    }

    function getCellSafe(cells, idx) {
      if (idx < 0) return '';
      var c = cells[idx];
      return (c && c.textContent ? c.textContent.toLowerCase() : '');
    }

    function filterTable() {
      var selectedItems = getSelectedValues(itemSelect);
      var selectedVendors = getSelectedValues(vendorSelect);
      var selectedBrands = getSelectedValues(brandSelect);
      var selectedBrandCats = getSelectedValues(brandCatSelect);
      var selectedDepts = getSelectedValues(deptSelect);
      var selectedPOL = getSelectedValues(polSelect);
      var selectedPRODUCT = getSelectedValues(productSelect);

      document.querySelectorAll('#excelTable tbody tr').forEach(function(row) {
        var cells = row.querySelectorAll('td');
        var item     = getCellSafe(cells, vizIdxFromCsv(${ITEM_INDEX}));
        var vendor   = getCellSafe(cells, vizIdxFromCsv(${VENDOR_INDEX + 1 + NEW_COL_SHIFT}));
        var brand    = getCellSafe(cells, vizIdxFromCsv(${BRAND_INDEX + 1 + NEW_COL_SHIFT}));
        var brandCat = getCellSafe(cells, vizIdxFromCsv(${BRAND_CATEGORY_INDEX + 1 + NEW_COL_SHIFT}));
        var dept     = getCellSafe(cells, vizIdxFromCsv(${DEPT_INDEX + 1 + NEW_COL_SHIFT}));
        var pol      = getCellSafe(cells, vizIdxFromCsv(${POL_INDEX + 1 + NEW_COL_SHIFT}));
        var product  = getCellSafe(cells, vizIdxFromCsv(${PRODUCT_INDEX + 1 + NEW_COL_SHIFT}));

        var show = (selectedItems.length === 0 || selectedItems.includes(item))
          && (selectedVendors.length === 0 || selectedVendors.includes(vendor))
          && (selectedBrands.length === 0 || selectedBrands.includes(brand))
          && (selectedBrandCats.length === 0 || selectedBrandCats.includes(brandCat))
          && (selectedDepts.length === 0 || selectedDepts.includes(dept))
          && (selectedPOL.length === 0 || selectedPOL.includes(pol))
          && (selectedPRODUCT.length === 0 || selectedPRODUCT.includes(product));

        row.style.display = show ? '' : 'none';
      });
    }

    [itemSelect, vendorSelect, brandSelect, brandCatSelect, deptSelect, polSelect, productSelect].forEach(function(el) {
      if (el) el.addEventListener('change', filterTable);
    });
  }

  // ═════════════════════════════════════════════════════════════
  //  CSV DOWNLOAD
  // ═════════════════════════════════════════════════════════════
  function initDownload() {
    var exportBtn = document.getElementById('downloadCsvBtn');
    if (exportBtn) {
      exportBtn.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        var url = (window.DOWNLOAD_URL || '').trim();
        if (url) { window.open(url, '_blank'); }
        else { alert('Download not available yet.'); }
      });
    }
  }

  // ═════════════════════════════════════════════════════════════
  //  SUBMIT HANDLER
  // ═════════════════════════════════════════════════════════════
  function initSubmit() {
    var suiteletForm = document.querySelector('form');
    var hiddenSelected = document.getElementById('custpage_selected');

    function collectSelectedRows() {
      if (!hiddenSelected) return;
      var out = [];
      document.querySelectorAll('#excelTable tbody tr').forEach(function(tr) {
        var cb = tr.querySelector('input[type="checkbox"][name^="row_select_"]');
        if (cb && cb.checked) {
          var rowId = (cb.name || '').split('_').pop();
          var qtyInput = tr.querySelector('input[type="number"][name^="qty_input_"]');
          var qty = qtyInput ? (qtyInput.value || '0') : '0';
          var memoInput = tr.querySelector('input[type="text"][name^="memo_input_"]');
          var memo = memoInput ? (memoInput.value || '') : '';
          var mosVal = '';
          var mosCell = tr.cells[2];
          if (mosCell) mosVal = (mosCell.textContent || '').trim();
          out.push({ rowId: rowId, qty: qty, memo: memo, mos: mosVal });
        }
      });
      hiddenSelected.value = JSON.stringify(out);
    }

    function closeHeaderPanel() {
      var open = document.querySelector('#filter-portal .th-filter-panel');
      if (open) open.remove();
      document.querySelectorAll('th.th-filter-active').forEach(function(th) {
        th.classList.remove('th-filter-active');
      });
    }

    function nativeSubmit() {
      if (!suiteletForm) return;
      collectSelectedRows();
      closeHeaderPanel();
      try {
        HTMLFormElement.prototype.submit.call(suiteletForm);
      } catch (e) {
        if (typeof suiteletForm.submit === 'function') suiteletForm.submit();
      }
    }

    var nsSubmitBtn = document.querySelector('input[type="submit"], button[type="submit"]');
    if (nsSubmitBtn) {
      nsSubmitBtn.addEventListener('click', function() {
        setTimeout(nativeSubmit, 0);
      }, true);
    }
    if (suiteletForm) {
      suiteletForm.addEventListener('submit', function() {
        setTimeout(nativeSubmit, 0);
      }, true);
    }
  }

  // ═════════════════════════════════════════════════════════════
  //  COLUMN HEADER FILTERS (Excel-style)
  // ═════════════════════════════════════════════════════════════
  function initColumnFilters() {
    var table = document.getElementById('excelTable');
    var container = document.getElementById('tableContainer');
    if (!table) return;

    var activeFilters = Object.create(null);

    function bodyRows() { return table.tBodies && table.tBodies[0] ? table.tBodies[0].rows : []; }

    function getCellText(row, idx) {
      var cells = row.cells;
      if (!cells || idx < 0 || idx >= cells.length) return '';
      return String(cells[idx].textContent || '').trim();
    }

    function normalizeVal(v) {
      var s = String(v == null ? '' : v).trim();
      if (/^-+\\s*none\\s*-+$/i.test(s)) s = '';
      if (/^nan$/i.test(s)) s = '';
      return s.toLowerCase();
    }

    // Returns ONLY values that exist in rows passing all OTHER column filters
    // This makes column B's dropdown update dynamically when column A is filtered
    function getAllValuesForColumn(colIdx) {
      var rows = bodyRows();
      var displayByKey = Object.create(null);
      for (var r = 0; r < rows.length; r++) {
        // Skip rows hidden by OTHER filters (not this column's filter)
        if (!rowPassesFiltersExceptColumn(rows[r], colIdx)) continue;
        var v = getCellText(rows[r], colIdx);
        if (v === '- None -' || v === 'NaN') v = '';
        if (v === '.00') v = '0.00';
        var key = normalizeVal(v);
        if (!displayByKey[key]) displayByKey[key] = (key === '' ? '(blank)' : (v || '(blank)'));
      }
      var keys = Object.keys(displayByKey).sort(function(a, b) {
        if (a === '' && b === '') return 0;
        if (a === '') return 1;
        if (b === '') return -1;
        return displayByKey[a].localeCompare(displayByKey[b], undefined, { numeric: true, sensitivity: 'base' });
      });
      return { keys: keys, displayByKey: displayByKey };
    }

    function rowPassesFiltersExceptColumn(row, exceptColIdx) {
      for (var k in activeFilters) {
        if (!Object.prototype.hasOwnProperty.call(activeFilters, k)) continue;
        var idx = parseInt(k, 10);
        if (idx === exceptColIdx) continue;
        var set = activeFilters[k];
        if (set && set.size > 0) {
          var v = normalizeVal(getCellText(row, idx));
          if (!set.has(v)) return false;
        }
      }
      return true;
    }

    function getValueCountsForColumn(colIdx) {
      var rows = bodyRows(), counts = Object.create(null);
      for (var r = 0; r < rows.length; r++) {
        if (!rowPassesFiltersExceptColumn(rows[r], colIdx)) continue;
        var v = getCellText(rows[r], colIdx);
        if (v === '- None -' || v === 'NaN') v = '';
        if (v === '.00') v = '0.00';
        var key = normalizeVal(v);
        counts[key] = (counts[key] || 0) + 1;
      }
      return counts;
    }

    function applyColumnFilters() {
      var rows = bodyRows(), pairs = [];
      for (var k in activeFilters) {
        if (!Object.prototype.hasOwnProperty.call(activeFilters, k)) continue;
        var s = activeFilters[k];
        if (s && s.size > 0) pairs.push([parseInt(k, 10), s]);
      }
      for (var r = 0; r < rows.length; r++) {
        var row = rows[r], show = true;
        for (var i = 0; i < pairs.length && show; i++) {
          var colIdx = pairs[i][0], set = pairs[i][1];
          var val = normalizeVal(getCellText(row, colIdx));
          if (!set.has(val)) show = false;
        }
        row.style.display = show ? '' : 'none';
      }
    }

    function openFilterPanel(btn, colIdx) {
      var th = btn.closest('th');
      var existing = document.querySelector('#filter-portal .th-filter-panel');
      if (existing) {
        var sameTh = existing.__ownerTH === th;
        existing.remove();
        if (existing.__teardown) existing.__teardown();
        th && th.classList.remove('th-filter-active');
        if (sameTh) return;
      }

      table.querySelectorAll('th.th-filter-active').forEach(function(node) {
        node.classList.remove('th-filter-active');
      });

      var panel = document.createElement('div');
      panel.className = 'th-filter-panel';
      panel.innerHTML =
        '<div class="hdr">Filter</div>' +
        '<input type="search" placeholder="Search values...">' +
        '<div class="list"></div>' +
        '<div class="actions">' +
        '<button type="button" data-act="clear">Clear</button>' +
        '<div style="display:flex;gap:6px;">' +
        '<button type="button" data-act="selectall">Select all</button>' +
        '<button type="button" data-act="deselectall">Deselect all</button>' +
        '<button type="button" data-act="apply">Apply</button>' +
        '</div></div>';

      var list = panel.querySelector('.list');
      panel.__ownerTH = th;

      var grouped = getAllValuesForColumn(colIdx);
      var universeKeys = grouped.keys;
      var displayByKey = grouped.displayByKey;
      var visibleCounts = getValueCountsForColumn(colIdx);
      var existingAllowed = activeFilters[colIdx] || null;

      var tempSelection = new Set();
      if (existingAllowed && existingAllowed.size > 0) {
        universeKeys.forEach(function(k) { if (existingAllowed.has(k)) tempSelection.add(k); });
      } else {
        universeKeys.forEach(function(k) { if ((visibleCounts[k] || 0) > 0) tempSelection.add(k); });
      }

      function renderList(filterText) {
        list.innerHTML = '';
        var ft = String(filterText || '').toLowerCase();

        // When user is actively searching, auto-select only matching items
        // and deselect non-matching ones (Excel-like behavior)
        if (ft) {
          universeKeys.forEach(function(k) {
            var labelText = displayByKey[k] || (k === '' ? '(blank)' : k);
            var matches = labelText.toLowerCase().indexOf(ft) !== -1;
            if (matches) {
              tempSelection.add(k);
            } else {
              tempSelection.delete(k);
            }
          });
        }

        universeKeys.forEach(function(k) {
          var labelText = displayByKey[k] || (k === '' ? '(blank)' : k);
          if (ft && labelText.toLowerCase().indexOf(ft) === -1) return;
          var id = 'f_' + colIdx + '_' + Math.random().toString(36).slice(2);
          var checked = tempSelection.has(k);
          var cnt = visibleCounts[k] || 0;
          var row = document.createElement('div');
          row.className = 'row';
          row.dataset.key = k;
          row.innerHTML =
            '<input type="checkbox" id="' + id + '" ' + (checked ? 'checked' : '') + '>' +
            '<label for="' + id + '">' + labelText + (cnt ? ' (' + cnt + ')' : '') + '</label>';
          list.appendChild(row);
        });
      }
      renderList('');

      // Save initial selection so we can restore when search is cleared
      var preSearchSelection = new Set(tempSelection);

      panel.querySelector('input[type="search"]').addEventListener('input', function() {
        var val = this.value;
        // When search is cleared, restore the pre-search selection state
        if (!val) {
          tempSelection.clear();
          preSearchSelection.forEach(function(k) { tempSelection.add(k); });
        }
        renderList(val);
      });

      list.addEventListener('change', function(e) {
        var cb = e.target;
        if (!cb || cb.type !== 'checkbox') return;
        var row = cb.closest('.row');
        if (!row) return;
        var key = row.dataset.key || '';
        if (cb.checked) tempSelection.add(key); else tempSelection.delete(key);
      });

      panel.querySelector('.actions').addEventListener('click', function(e) {
        var act = e.target && e.target.getAttribute('data-act');
        if (!act) return;

        if (act === 'clear') {
          delete activeFilters[colIdx];
          th.classList.remove('th-filter-active');
          panel.remove();
          if (panel.__teardown) panel.__teardown();
          applyColumnFilters();
          return;
        }
        if (act === 'selectall') {
          list.querySelectorAll('input[type="checkbox"]').forEach(function(cb) {
            cb.checked = true;
            tempSelection.add(cb.closest('.row').dataset.key || '');
          });
          return;
        }
        if (act === 'deselectall') {
          list.querySelectorAll('input[type="checkbox"]').forEach(function(cb) {
            cb.checked = false;
            tempSelection.delete(cb.closest('.row').dataset.key || '');
          });
          return;
        }
        if (act === 'apply') {
          var allow = new Set();
          tempSelection.forEach(function(k) { allow.add(k); });
          if (allow.size === universeKeys.length) {
            delete activeFilters[colIdx];
            th.classList.remove('th-filter-active');
          } else {
            activeFilters[colIdx] = allow;
            th.classList.add('th-filter-active');
          }
          panel.remove();
          if (panel.__teardown) panel.__teardown();
          applyColumnFilters();
        }
      });

      var portal = document.getElementById('filter-portal') || document.body;
      portal.appendChild(panel);

      function positionPanel() {
        var rect = btn.getBoundingClientRect();
        var vw = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
        var ph = panel.offsetHeight || 0;
        var pw = panel.offsetWidth || 270;
        var belowTop = rect.bottom + 6;
        var aboveTop = rect.top - ph - 6;
        var top = (belowTop + ph <= window.innerHeight) ? belowTop : Math.max(8, aboveTop);
        var left = Math.min(vw - pw - 8, Math.max(8, rect.right - pw));
        panel.style.top = Math.round(top) + 'px';
        panel.style.left = Math.round(left) + 'px';
      }
      positionPanel();

      var posRaf = false;
      function schedulePosition() {
        if (posRaf) return;
        posRaf = true;
        requestAnimationFrame(function() { posRaf = false; positionPanel(); });
      }

      window.addEventListener('scroll', schedulePosition, true);
      window.addEventListener('resize', schedulePosition);
      if (container) container.addEventListener('scroll', schedulePosition);

      function teardown() {
        window.removeEventListener('scroll', schedulePosition, true);
        window.removeEventListener('resize', schedulePosition);
        if (container) container.removeEventListener('scroll', schedulePosition);
      }
      panel.__teardown = teardown;

      setTimeout(function() {
        document.addEventListener('click', function handler(e) {
          if (panel.contains(e.target) || th.contains(e.target)) return;
          panel.remove();
          th.classList.remove('th-filter-active');
          teardown();
          document.removeEventListener('click', handler);
        });
      }, 0);

      th.classList.add('th-filter-active');
    }

    // Delegate header clicks
    if (table.tHead) {
      table.tHead.addEventListener('click', function(e) {
        var btn = e.target.closest('.th-filter-btn');
        if (!btn) return;
        var colIdx = parseInt(btn.getAttribute('data-col'), 10);
        if (!isFinite(colIdx)) return;
        e.stopPropagation();
        openFilterPanel(btn, colIdx);
      });
    }
  }

  // ═════════════════════════════════════════════════════════════
  //  STICKY COLUMNS & SCROLLBAR
  // ═════════════════════════════════════════════════════════════
  function initTableLayout() {
    var container = document.getElementById('tableContainer');
    var table = document.getElementById('excelTable');
    var topScroll = document.getElementById('topScroll');
    var topInner = document.getElementById('topScrollInner');
    if (!table) return;

    var widthsApplied = false;
    var stickyApplied = false;

    var FIXED_STICKY_WIDTHS = [50, 100, 110, 110, 90, 200, 110, 180, 90, 450];

    function applyFixedWidths(widths) {
      if (!table || !table.tHead || !table.tBodies[0]) return;
      var ths = table.tHead.rows[0].cells;
      for (var i = 0; i < widths.length && i < ths.length; i++) {
        ths[i].style.width = widths[i] + 'px';
        ths[i].style.minWidth = widths[i] + 'px';
        ths[i].style.maxWidth = widths[i] + 'px';
      }
      if (widthsApplied) return;
      widthsApplied = true;
      var rows = table.tBodies[0].rows;
      for (var r = 0; r < rows.length; r++) {
        var tds = rows[r].cells;
        for (var c = 0; c < widths.length && c < tds.length; c++) {
          tds[c].style.width = widths[c] + 'px';
          tds[c].style.minWidth = widths[c] + 'px';
          tds[c].style.maxWidth = widths[c] + 'px';
          tds[c].style.overflow = 'hidden';
          tds[c].style.textOverflow = 'ellipsis';
        }
      }
    }

    function updateTopScrollbarWidth() {
      if (!container || !table || !topInner) return;
      topInner.style.width = Math.max((table.scrollWidth || 0), (container.clientWidth || 0) + 2) + 'px';
    }

    if (topScroll) topScroll.style.display = 'block';
    if (topScroll && container) {
      topScroll.addEventListener('scroll', function() { container.scrollLeft = topScroll.scrollLeft; });
      container.addEventListener('scroll', function() { topScroll.scrollLeft = container.scrollLeft; });
    }

    function makeSticky(firstN) {
      if (!table || !table.tHead || !table.tHead.rows.length) return;
      var lefts = [];
      var acc = 0;
      for (var i = 0; i < firstN; i++) {
        var w = FIXED_STICKY_WIDTHS[i];
        if (typeof w !== 'number' || isNaN(w)) {
          var hc = table.tHead.rows[0].cells[i];
          w = hc ? Math.round(hc.getBoundingClientRect().width) : 100;
        }
        lefts[i] = acc;
        acc += w;
      }
      var allRows = table.rows;
      for (var r = 0; r < allRows.length; r++) {
        var cells = allRows[r].cells;
        for (var c = 0; c < firstN && c < cells.length; c++) {
          var cell = cells[c];
          if (!stickyApplied) {
            cell.classList.add('sticky-col');
            if (c > 0) cell.classList.add('sep-left');
          }
          cell.style.left = lefts[c] + 'px';
          cell.style.zIndex = (r === 0 ? 30 : 20) + c;
        }
      }
      stickyApplied = true;
    }

    function init() {
      updateTopScrollbarWidth();
      applyFixedWidths(FIXED_STICKY_WIDTHS);
      makeSticky(11);
      if (topScroll && container) topScroll.scrollLeft = container.scrollLeft;
    }

    requestAnimationFrame(init);
    setTimeout(init, 200);

    var resizeTimer = null;
    window.addEventListener('resize', function() {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(init, 180);
    });
  }

  // ═════════════════════════════════════════════════════════════
  //  BOOT
  // ═════════════════════════════════════════════════════════════
  document.addEventListener('DOMContentLoaded', function() {
    initRowEvents();
    initTopFilters();
    initDownload();
    initSubmit();
    initColumnFilters();
    initTableLayout();
  });

})();
</script>
`;

        // Save cleaned file
        var newfileObj = file.create({
          name: "Download_" + fileObj.name,
          fileType: file.Type.CSV,
          contents: newContent.map(row => row.join(",")).join("\n"),
          encoding: file.Encoding.UTF8,
          folder: 378271,
          isOnline: true
        });
        var newFileId = newfileObj.save();
        fileField.defaultValue = newFileId;

        var newfileObj2 = file.create({
          name: "RR Tool Details.csv",
          fileType: file.Type.CSV,
          contents: newContent.map(row => row.join(",")).join("\n"),
          encoding: file.Encoding.UTF8,
          folder: 279208,
          isOnline: true
        });
        var newFileId2 = newfileObj2.save();

        var reloadFile = file.load({ id: newFileId2 });
        var dlUrl = reloadFile.url;

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

        html += '<script>window.DOWNLOAD_URL = ' + JSON.stringify(dlUrl || '') + ';</script>';
        html += '<script>'
          + 'window.__ADMIN_IDX__=' + JSON.stringify(adminCsvIndex) + ';'
          + 'window.__TRUNC_AFTER__=' + JSON.stringify(truncateAfterAdmin && adminCsvIndex >= 0) + ';'
          + 'window.__ADMIN_REMOVED__=' + JSON.stringify(removeJustAdmin && adminCsvIndex >= 0) + ';'
          + '</script>';

        // Close the rr-root div
        html += '</div>';

        htmlField.defaultValue = html;
        // NO client script module path — everything is inline now
        context.response.writePage(form);
      }

      if (context.request.method === 'POST') {
        var params = context.request.parameters;
        var fileId = params.custpage_file_id;

        var p = context.request.parameters || {};
        var postedEmp = p.custpage_empid || '';
        var postedTs = p.custpage_ts || '';
        var postedSig = p.custpage_sig || '';

        var authorized = false;
        if (postedEmp && postedTs && postedSig) {
          authorized = verify(postedEmp, postedTs, postedSig);
        }

        if (!authorized) {
          context.response.write(`
            <html><head>
            <script>setTimeout(function(){ window.location.href = ${JSON.stringify(portalUrl)}; }, 1200);</script>
            <style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>
            </head><body><div class="message">Session expired. Please log in again.</div></body></html>
          `);
          return;
        }

        // Load the file and parse contents
        var fileObj = file.load({ id: fileId });
        var content = fileObj.getContents();
        var rows = content.split('\n');
        var selectedJson = params.custpage_selected;
        var selectedRows = [];
        var createdCount = 0;

        if (selectedJson) {
          try {
            var picked = JSON.parse(selectedJson);
            picked.forEach(function (p) {
              var rowId = String(p.rowId || '').trim();
              var qty = parseInt(p.qty, 10) || 0;
              var memo = p.memo || 0;
              var mos = parseFloat(p.mos);
              if (!rowId) return;
              var columns = rows[rowId].split(/,(?=(?:[^\"]*"[^"]*")*[^\"]*$)/)
                .map(function (val) { return val.replace(/"/g, '').trim(); });
              selectedRows.push({ rowId: rowId, qty: qty, memo: memo, monthStock: mos, columns: columns });
            });
          } catch (e) {
            log.error('Bad custpage_selected JSON', e);
          }
        }

        // Fallback to legacy scanning if nothing captured
        if (selectedRows.length === 0) {
          for (var key in params) {
            if (key.startsWith('row_select_')) {
              var rowId = key.split('_')[2];
              log.debug('params', params);
              var qty = parseInt(params['qty_input_' + rowId]) || 0;
              var memo = params['memo_input_' + rowId] || '';
              var columns = rows[rowId].split(/,(?=(?:[^\"]*"[^"]*")*[^\"]*$)/)
                .map(function (val) { return val.replace(/"/g, '').trim(); });
              selectedRows.push({ rowId: rowId, qty: qty, memo: memo, columns: columns });
            }
          }
        }

        log.debug('selectedRows', selectedRows);

        selectedRows.forEach(function (entry) {
          var cols = entry.columns;
          log.debug('cols', cols);

          try {
            record.create({
              type: 'customrecord_mi_planned_po',
              isDynamic: true
            })
              .setValue({ fieldId: 'custrecord_mi_item', value: cols[3] })
              .setValue({ fieldId: 'custrecord_mi_order_qty', value: entry.qty })
              .setValue({ fieldId: 'custrecord_mi_purchase_memo', value: entry.memo == 0 ? '' : entry.memo })
              .setValue({ fieldId: 'custrecord_month_of_stocks', value: safeParseFloat(entry.monthStock) })
              .setValue({ fieldId: 'custrecord_mi_qty_of_ordered_not_ship', value: safeParseFloat(cols[26]) })
              .setValue({ fieldId: 'custrecord_mi_qty_available', value: safeParseFloat(cols[33]) })
              .setValue({ fieldId: 'custrecord_mi_qty_in_transit', value: safeParseFloat(cols[27]) })
              .setValue({ fieldId: 'custrecord_mi_min_month_qty', value: safeParseFloat(cols[18]) })
              .save();

            createdCount++;
          } catch (e) {
            log.error('Error creating custom record for row ' + entry.rowId, e);
          }
        });

        context.response.write(`
          <html>
          <head>
          <meta http-equiv="refresh" content="5;URL=https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2108&deploy=1&compid=4975346&ns-at=AAEJ7tMQmJxVsovhMpsEMUF39xnBuyMwWM4G2T7SnvA62twq8hg&empid=${encodeURIComponent(postedEmp)}&ts=${encodeURIComponent(postedTs)}&sig=${encodeURIComponent(postedSig)}" />
          <style>
            body {
              font-family: 'DM Sans', Arial, sans-serif;
              text-align: center;
              padding-top: 100px;
              background: #f8fafc;
            }
            .message-box {
              display: inline-block;
              background: #fff;
              padding: 24px 36px;
              border: 1px solid #e2e8f0;
              border-radius: 12px;
              color: #1e293b;
              font-size: 16px;
              box-shadow: 0 4px 12px rgba(0,0,0,0.06);
            }
            .message-box strong { color: #2563eb; font-size: 20px; }
          </style>
          </head>
          <body>
            <div class="message-box">
              <strong>${createdCount}</strong> Planned Purchase Order(s) created.<br />
              You will be redirected to the main page in 5 seconds...
            </div>
          </body>
          </html>
        `);
      }
    }

    function getInventoryBalanceMap() {
      var resultMap = {};

      var inventorybalanceSearchObj = search.create({
        type: "inventorybalance",
        filters: [
          ["status", "anyof", "6", "1", "3", "5", "8"],
          "AND",
          ["item.isinactive", "is", "F"]
        ],
        columns: [
          search.createColumn({ name: "item", summary: "GROUP" }),
          search.createColumn({ name: "onhand", summary: "SUM" }),
          search.createColumn({
            name: "formulanumeric", summary: "SUM",
            formula: "case when {status} = 'Good' then {onhand} else 0 end"
          }),
          search.createColumn({
            name: "formulanumeric1", summary: "SUM",
            formula: "case when {status} = 'Deviation' then {onhand} else 0 end"
          }),
          search.createColumn({
            name: "formulanumeric2", summary: "SUM",
            formula: "case when {status} = 'Hold' then {onhand} else 0 end"
          }),
          search.createColumn({
            name: "formulanumeric3", summary: "SUM",
            formula: "case when {status} = 'Inspection' then {onhand} else 0 end"
          }),
          search.createColumn({
            name: "formulanumeric4", summary: "SUM",
            formula: "case when {status} = 'Label' then {onhand} else 0 end"
          }),
          search.createColumn({ name: "available", summary: "SUM", label: "Available" })
        ]
      });

      inventorybalanceSearchObj.run().each(function (result) {
        var itemId = result.getValue({ name: "item", summary: "GROUP" });
        var onH = result.getValue({ name: "onhand", summary: "SUM" });
        var goodQty = parseFloat(result.getValue({ name: "formulanumeric", summary: "SUM" })) || 0;
        var badQty = parseFloat(result.getValue({ name: "formulanumeric1", summary: "SUM" })) || 0;
        var holdQty = parseFloat(result.getValue({ name: "formulanumeric2", summary: "SUM" })) || 0;
        var inspectQty = parseFloat(result.getValue({ name: "formulanumeric3", summary: "SUM" })) || 0;
        var labelQty = parseFloat(result.getValue({ name: "formulanumeric4", summary: "SUM" })) || 0;
        var total = parseFloat((goodQty + badQty).toFixed(2));
        var avail = parseFloat(result.getValue({ name: "available", summary: "SUM" })) || 0;

        resultMap[itemId] = { good: goodQty, bad: badQty, hold: holdQty, inspect: inspectQty, label: labelQty, total: total, avail: avail, onH: onH };
        return true;
      });

      return resultMap;
    }

    function getPurchaseOrderQtyMap() {
      var resultMap = {};

      var purchaseorderSearchObj = search.create({
        type: "transaction",
        settings: [{ "name": "consolidationtype", "value": "ACCTTYPE" }],
        filters: [
          ["type", "anyof", "VendBill","VendCred"],
          // "AND",
          // ["closed", "is", "F"],
          "AND",
          ["mainline", "is", "F"],
          // "AND",
          // ["cogs", "is", "F"],
          "AND",
          ["shipping", "is", "F"],
          "AND",
          ["item", "noneof", "@NONE@"],
          "AND",
          [["trandate", "within", "lastyear"], "OR", ["trandate", "within", "thisyear"]]
        ],
        columns: [
          search.createColumn({ name: "item", summary: "GROUP", label: "Item" }),
          search.createColumn({
            name: "formulanumeric", summary: "SUM",
            formula: "CASE WHEN TO_CHAR({trandate}, 'YYYY') = TO_CHAR(ADD_MONTHS(SYSDATE, -12), 'YYYY') THEN ABS({quantity}) ELSE 0 END",
            label: "Last Year"
          }),
          search.createColumn({
            name: "formulanumeric", summary: "SUM",
            formula: "CASE WHEN TO_CHAR({trandate}, 'YYYY') = TO_CHAR(SYSDATE, 'YYYY') THEN ABS({quantity}) ELSE 0 END",
            label: "This Year"
          })
        ]
      });

      purchaseorderSearchObj.run().each(function (result) {
        var itemId = result.getValue({ name: "item", summary: "GROUP" });
        var cols = result.columns;
        var lastYear = parseFloat(result.getValue(cols[1])) || 0;
        var thisYear = parseFloat(result.getValue(cols[2])) || 0;

        resultMap[itemId] = { lastYear: lastYear, thisYear: thisYear };
        return true;
      });

      log.debug('PO Qty Map count', Object.keys(resultMap).length);
      return resultMap;
    }

    function getCommBackMap() {
      var resultMap = {};

      var itemSearchObj = search.create({
    type: "item",
    filters: [],
    columns: [
        search.createColumn({
            name: "internalid",
            summary: "GROUP",
            sort: search.Sort.ASC,
            label: "Internal ID"
        }),
        search.createColumn({
            name: "formulanumeric",
            summary: "SUM",
            formula: "NVL({locationquantitycommitted},0) + NVL({locationtoresvcommitted},0)",
            label: "Formula (Numeric)"
        }),
        search.createColumn({
            name: "locationquantitybackordered",
            summary: "SUM",
            label: "Location Back Ordered"
        })
    ]
});

var pagedData = itemSearchObj.runPaged({
    pageSize: 1000
});

pagedData.pageRanges.forEach(function (pageRange) {
    var page = pagedData.fetch({ index: pageRange.index });

    page.data.forEach(function (result) {
        var itemId = result.getValue({
            name: "internalid",
            summary: "GROUP"
        });

        var qtyComm = result.getValue({
            name: "formulanumeric",
            summary: "SUM"
        });

        var qtyBack = result.getValue({
            name: "locationquantitybackordered",
            summary: "SUM"
        });

        resultMap[itemId] = {
            qtyComm: qtyComm,
            qtyBack: qtyBack
        };
    });
});

      return resultMap;
    }

    function safeParseFloat(val) {
      var num = parseFloat(val);
      return isNaN(num) ? 0 : num;
    }
    function safeParseInt(val) {
      var num = parseInt(val);
      return isNaN(num) ? 0 : num;
    }

    return {
      onRequest: onRequest
    };
  });