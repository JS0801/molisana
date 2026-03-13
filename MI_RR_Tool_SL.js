/**
* @NApiVersion 2.1
* @NScriptType Suitelet
*/
define(['N/ui/serverWidget', 'N/file', 'N/log', 'N/search', 'N/record', 'N/runtime', 'N/crypto'],
function(ui, file, log, search, record, runtime, crypto) {
  function onRequest(context) {
    
    var portalUrl = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2110&deploy=1&compid=4975346&ns-at=AAEJ7tMQamzukv1WMqTK6i2c27bRetbrd2MDLjhDgPPFOawMxCo';
    
    // --- Signed session helpers (same as Inventory Dashboard) ---
    const SECRET = runtime.getCurrentScript().getParameter({ name: 'custscript_portal_secret' }) || 'change-me';//'rR9Z7KpXw2N6C8mE4HqFJYvT5bS0aUeD1LQG3oM';
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
      if (!empid || !ts || !sig) return
      false;
      if (Math.abs(Date.now() - parseInt(ts, 10)) > TOKEN_TTL_MS) return false;
      try { return sign(empid, ts) === sig; } catch (e) { log.error('verify token', e); return false; }
    }
    
    if (context.request.method === 'GET') {
      
      var q = context.request.parameters || {};
      var typeParam = (q.type || 3).toString();
      var showTopFilters = (typeParam !== '4');
      var formName = 'MI Reorder Tool (' + (typeParam == 4? 'Basic': 'Admin') + ')';
      var mode = (q.mode || '').toString().toLowerCase(); // 'cron' or ''
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
      
      
      
      // If a valid signature is present, prefer that; else keep legacy selectedEmp
      if (empid && ts && sig && verify(empid, ts, sig)) {
        selectedEmp = empid;
      }
      // If neither signature nor legacy emp id, bounce to portal
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
        
        folderSearchObj.run().each(function(result) {
          log.debug('result', result)
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
        log.debug('rows', rows.length)
        
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
        log.debug('rows', rows.length)
        
        var newContent = [];
        const ITEM_INDEX = 2;          // already used
        const VENDOR_INDEX = 35;       // already used
        const BRAND_CATEGORY_INDEX = 33;
        const BRAND_INDEX = 32;
        const DEPT_INDEX = 52;
        const POL_INDEX = 36;
        const PRODUCT_INDEX = 53;
        
        const uniqueItems = new Set();
        const uniqueVendors = new Set();
        const uniqueBrands = new Set();
        const uniqueBrandCategories = new Set();
        const uniqueDepartments = new Set();
        const uniquePOL = new Set();
        const uniquePRODUCT = new Set();
        
        let html = `
        <style>
        
        /* ---- Column filter UI ---- */
        .th-filter-wrap{ position:relative; display:inline-flex; align-items:center; gap:6px; }
        .th-filter-btn{
          cursor:pointer; border:1px solid #cbd5e1; background:#fff; padding:2px 4px; border-radius:4px;
          line-height:1; font-size:11px; user-select:none;
        }
        .th-filter-btn:hover{ background:#f3f4f6; }
        .th-filter-panel{
          position: fixed;                 /* was: absolute */
          top: 0; left: 0;                 /* will be set via JS */
          width: 260px; max-height: 320px; overflow: auto;
          background: #fff; border:1px solid #cbd5e1; border-radius:8px; padding:10px; margin-top:0;
          box-shadow:0 8px 24px rgba(0,0,0,.12);
          z-index: 9999;                   /* above sticky cells & everything else */
          pointer-events: auto;
        }
        
        .th-filter-panel .hdr{ font-weight:600; font-size:12px; margin-bottom:6px; }
        .th-filter-panel input[type="search"]{ width:100%; padding:6px 8px; box-sizing:border-box; margin-bottom:8px; }
        .th-filter-panel .list{ max-height:200px; overflow:auto; border:1px solid #e5e7eb; border-radius:6px; padding:6px; }
        .th-filter-panel .row{ display:flex; align-items:center; gap:8px; padding:2px 0; font-size:12px; }
        .th-filter-panel .actions{ display:flex; justify-content:space-between; gap:8px; margin-top:8px; }
        .th-filter-panel .actions button{
          padding: 2px 6px;        /* smaller */
          font-size: 11px;         /* smaller */
          line-height: 1.2;
          border-radius: 4px;
          border: 1px solid #cbd5e1;
          background: #fff;
          cursor: pointer;
          height: auto;
          min-height: 22px;        /* keeps them clickable but compact */
        }
        .th-filter-panel .actions button[data-act="apply"]{
          font-weight: 600;
          border-color: #9ab3ff;
        }
        
        /* optional subtle danger for Clear */
        .th-filter-panel .actions button[data-act="clear"]{
          color: #7f1d1d;
          border-color: #f3d2d2;
        }
        .th-filter-active .th-filter-btn{ border-color:#2563eb; background:#eff6ff; }
        
        
        .table-container{
          max-height: 850px;
          overflow-y: auto;
          overflow-x: auto;
          border: 1px solid #ccc;
        }
        
        /* TOP sync scrollbar */
        .h-scroll {
          height: 16px;
          overflow-x: auto;
          overflow-y: hidden;
          border: 1px solid #ccc;
          border-bottom: 0;
          width: 100%;           /* <-- add this */
        }
        .h-scroll-inner { height: 1px; }
        .table-viewport{
          width: 100%;
          max-width: 100%;
          box-sizing: border-box;
        }
        
        /* Force a real viewport cap so the table can overflow horizontally */
        .ns-table-viewport-cap{
          /* pick a sensible cap smaller than your table width */
          max-width: calc(100vw - 10px);
          width: 100%;
          margin: 0;
          padding: 0;
        }
        
        /* Make the scroller and table area respect the viewport width */
        .h-scroll{ width: 100%; }     /* you already have this */
        .table-container{
          width: 100%;
          max-width: 100%;
          overflow-x: auto;           /* horizontal scroll lives here */
          overflow-y: auto;
        }
        
        /* Table */
        #excelTable{
          border-collapse: separate;
          width: max-content;
          min-width: 100%;
          table-layout: auto;
        }
        
        #excelTable th,#excelTable td{
          border: 1px solid #ccc;
          padding: 8px 12px;
          background-color: #fff;
          white-space: nowrap;
          font-size: 12px;
        }
        
        /* Sticky HEADER */
        #excelTable thead th{
          position: sticky;
          top: 0;
          background-color: #f3f3f3;
          z-index: 9;
        }
        
        /* Sticky COLUMNS (1 & 2) */
        #excelTable th.col-sticky-1, #excelTable td.col-sticky-1{
          position: sticky;
          left: 0;
          background: #f9f9f9;
          z-index: 10; /* above thead background */
          min-width: 50px;
        }
        #excelTable th.col-sticky-2, #excelTable td.col-sticky-2{
          position: sticky;
          left: var(--sticky-left-1, 120px); /* set by JS based on actual width of col 1 */
          background: #f9f9f9;
          z-index: 10;
          min-width: 120px;
        }
        
        /* Inputs */
        input[type="number"]{
          width: 100%;
          box-sizing: border-box;
          padding: 4px;
        }
        input[type="checkbox"]{ transform: scale(1.2); }
        
        #excelTable thead th{
          position: sticky;
          top: 0;
          background-color: #f3f3f3;
          z-index: 9;
        }
        
        /* Sticky columns base */
        #excelTable th.col-sticky-1, #excelTable td.col-sticky-1{
          position: sticky;
          left: 0;
          background: #f9f9f9;
          z-index: 10;
          min-width: 50px;
        }
        #excelTable th.col-sticky-2, #excelTable td.col-sticky-2{
          position: sticky;
          left: var(--sticky-left-1, 120px);
          background: #f9f9f9;
          z-index: 10;
          min-width: 120px;
        }
        
        /* Ensure header cells that are also sticky columns sit on top */
        #excelTable thead th.col-sticky-1,
        #excelTable thead th.col-sticky-2{
          z-index: 12;
        }
        </style>
        
        
        <style>
        .download-btn {
          background-color: white;
          color: white;
          border: none;
          font-size: 13px;
          cursor: pointer;
          transition: background-color 0.2s ease;
        }
        
        .download-btn:hover {
          background-color: #005f8d;
        }
        
        .filter-hover-wrapper {
          position: relative;
          display: inline-block;
          margin-bottom: 15px;
          z-index: 4000;
        }
        
        /* Generic sticky column styling */
        .sticky-col{
          position: sticky;
          background: #f9f9f9;    /* so it covers cells when scrolling */
          background-clip: padding-box;
          z-index: 11;
        }
        thead th.sticky-col{
          z-index: 30;            /* header above body cells */
        }
        .sticky-col.sep-left{
          border-left-color: transparent;   /* avoid double border where columns meet */
          box-shadow: inset 1px 0 0 #ccc;   /* crisp left divider that doesn't jitter */
        }
        .filter-button {
          color: #0073aa; /* NetSuite-like action blue */
          font-size: 15px;
          font-weight: 600;
          cursor: pointer;
          display: inline-block;
          padding: 4px 8px;
          border-radius: 4px;
          transition: all 0.2s ease;
          background-color: transparent;
          border: none;
        }
        
        .filter-button:hover {
          color: #005f8d;
          background-color: #eef7ff;
          text-decoration: underline;
        }
        
        
        .filter-panel {
          display: none;
          position: absolute;
          top: 110%;
          left: 0;
          background-color: white;
          border: 1px solid #ccc;
          padding: 15px;
          border-radius: 6px;
          width: 650px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.15);
          z-index: 5000;
        }
        
        .filter-grid {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 15px;
        }
        
        .filter-grid label {
          font-weight: 600;
          font-size: 12px;
          display: block;
          margin-bottom: 5px;
        }
        
        .filter-grid select {
          width: 100%;
        }
        .top-controls{
          display:flex;
          align-items:center;
          justify-content:space-between;
          gap:12px;
          margin-bottom:10px;
        }
        
        .controls-right{
          display:flex;
          align-items:center;
          gap:12px;
        }
        .totals-bar{
          display:inline-flex;
          align-items:center;
          gap:8px;
          padding:6px 12px;
          border:1px solid #a3d2f2;
          border-radius:999px;
          background:#f3f9ff;
          box-shadow:0 1px 2px rgba(0,0,0,0.04);
        }
        
        .totals-label{
          font-size:12px;
          color:#005b99;
          font-weight:600;
          letter-spacing:.2px;
          text-transform:uppercase;
        }
        
        .totals-value{
          font-feature-settings:"tnum";
          font-variant-numeric:tabular-nums;
          font-weight:700;
          font-size:14px;
          color:#003f6b;
        }
        </style>
        <div style="margin-bottom: 10px;">
        <button id="downloadCsvBtn" class="download-btn" type="button"><img src="https://cdn-icons-png.flaticon.com/512/10630/10630240.png"
        alt="csv-icon"
        style="width: 16px; vertical-align: middle; margin-right: 6px;" /></button>
        ${showTopFilters ? `
          <div class="filter-hover-wrapper" id="filterWrapper">
          <div class="filter-button" id="filterToggle">Filters</div>
          <div class="filter-panel" id="filterPanel">
          <div class="filter-grid">
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
          ` : ``}
          </div>
          <div id="totalsBar" class="totals-bar" title="Sum of Item Cubic Space for all selected rows"> <span class="totals-label">Total Cubic Space</span> <span id="totalCubicValue" class="totals-value">0</span> </div> <div id="totalsBarWgt" class="totals-bar" title="Sum of Item Weight for all selected rows"> <span class="totals-label">Total Weight</span> <span id="totalWeightValue" class="totals-value">0</span> </div>
          
          <script>
          document.addEventListener('DOMContentLoaded', function () {
            const wrapper = document.getElementById('filterWrapper');
            const panel = document.getElementById('filterPanel');
            
            let timer;
            
            wrapper.addEventListener('mouseenter', function () {
              clearTimeout(timer);
              panel.style.display = 'block';
            });
            
            wrapper.addEventListener('mouseleave', function () {
              timer = setTimeout(() => {
                panel.style.display = 'none';
              }, 200); // Small delay to allow smooth exit
            });
          });
          </script>
          <div class="table-viewport">
          <div class="ns-table-viewport-cap">
          <div class="h-scroll" id="topScroll">
          <div class="h-scroll-inner" id="topScrollInner"></div>
          </div>
          <div class="table-container">
          <table id="excelTable">
          
          <thead><tr>
          `;
          
          const headers = rows[0].split(/,(?=(?:[^\"]*"[^"]*")*[^\"]*$)/);
          
          function norm(h){
            return String(h || '')
            .replace(/^[\uFEFF\s"]+|[\s"]+$/g, '')  // trim + strip quotes/BOM
            .replace(/\s+/g, ' ')
            .toLowerCase();
          }
          
          const rawHeaders = headers;                  // already split with your CSV regex
          const normalized = rawHeaders.map(norm);
          
          let adminCsvIndex = normalized.indexOf('admin portal');
          if (adminCsvIndex < 0) {
            adminCsvIndex = normalized.findIndex(h =>
              h === 'admin' || h.startsWith('admin portal') || h.includes('admin portal')
            );
          }
          
          // Behaviour flags
          const truncateAfterAdmin = (typeParam === '4'); // remove Admin and everything after
          const removeJustAdmin    = (typeParam === '3'); // remove only Admin
          
          // Headers to SHOW & DOWNLOAD (after removal/truncation)
          const headersForOutput = (() => {
            if (adminCsvIndex >= 0) {
              if (truncateAfterAdmin) return rawHeaders.slice(0, adminCsvIndex);
              if (removeJustAdmin)    return rawHeaders.filter((_, i) => i !== adminCsvIndex);
            }
            return rawHeaders.slice();
          })();
          headersForOutput.push('Filter');
          newContent.push(headersForOutput);
          
          
          
          
          html += '<th class="col-sticky-1">Select</th>'
          + '<th class="col-sticky-2">Order Qty</th>'
          + '<th class="col-sticky-2">Month of Stock</th>'
          + '<th class="col-sticky-2">Item Space</th>'
          + '<th class="col-sticky-2">Weight</th>';
          
          var dataColStart = 5; // 5 control columns before CSV data
          var visibleIdx = 0;
          
          headersForOutput.forEach(function (value, outIdx) {
            // Map outIdx back to original CSV idx (needed for filters later)
            // If we removed Admin col, any original idx >= adminCsvIndex shifts by -1 in headersForOutput.
            // But for the DOM header (data-col), we only need the DOM index:
            var domIndex = dataColStart + visibleIdx;
            visibleIdx++;
            
            value = String(value || '').replace(/"/g,'').trim();
            var stickyCls = (outIdx < 5 ? ' col-sticky-2' : '');
            var safeText = value || ('Col ' + (outIdx + 1));
            
            html += '<th class="' + (outIdx < 5 ? 'col-sticky-2' : '') + '" data-domidx="'+domIndex+'">'
            +   '<span class="th-filter-wrap">'
            +     '<span class="th-label">'+ safeText +'</span>'
            +     '<button class="th-filter-btn" type="button" data-col="'+ domIndex +'">▾</button>'
            +   '</span>'
            + '</th>';
          });
          
          html += '</tr></thead><tbody>';
          
          let redRows = [];
          let blackRows = [];
          
          function csvIndexIsExposed(idx){
            if (idx == adminCsvIndex || idx > adminCsvIndex) idx = idx + 1;
            return idx;
          }
          
          var balances = getInventoryBalanceMap();
          log.debug('Inventory Balance Map', balances);
          var alreadyExsist = {};
          rows.slice(1).forEach(function (row) {
            if (!row || row.trim() === '') return;
            
            const columns = row.split(/,(?=(?:[^\"]*"[^"]*")*[^\"]*$)/);
            if (!columns.length) return;
            
            // --- Filter (type=4) on the ORIGINAL row BEFORE any changes:
            if (typeParam === '4' && adminCsvIndex >= 0) {
              var adminCellVal = (columns[adminCsvIndex] || '').replace(/"/g,'').trim().toLowerCase();
              if (adminCellVal !== 'admin portal') return; // skip non-Admin rows
            }
            
            let calcCols = columns.slice();
            const monthAvg = parseFloat(calcCols[csvIndexIsExposed(11)]);
            const itemid = calcCols[csvIndexIsExposed(3)];
            
            if (alreadyExsist.hasOwnProperty(itemid) || !itemid) return;
            alreadyExsist[itemid] = true;
            
            if (columns[csvIndexIsExposed(ITEM_INDEX)])   uniqueItems.add(columns[csvIndexIsExposed(ITEM_INDEX)].replace(/"/g,'').trim());
            if (columns[csvIndexIsExposed(VENDOR_INDEX)]) uniqueVendors.add(columns[csvIndexIsExposed(VENDOR_INDEX)].replace(/"/g,'').trim());
            if (columns[csvIndexIsExposed(BRAND_INDEX)])  uniqueBrands.add(columns[csvIndexIsExposed(BRAND_INDEX)].replace(/"/g,'').trim());
            if (columns[csvIndexIsExposed(BRAND_CATEGORY_INDEX)]) uniqueBrandCategories.add(columns[csvIndexIsExposed(BRAND_CATEGORY_INDEX)].replace(/"/g,'').trim());
            if (columns[csvIndexIsExposed(DEPT_INDEX)])   uniqueDepartments.add(columns[csvIndexIsExposed(DEPT_INDEX)].replace(/"/g,'').trim());
            if (columns[csvIndexIsExposed(POL_INDEX)])    uniquePOL.add(columns[csvIndexIsExposed(POL_INDEX)].replace(/"/g,'').trim());
            if (columns[csvIndexIsExposed(PRODUCT_INDEX)])    uniquePRODUCT.add(columns[csvIndexIsExposed(PRODUCT_INDEX)].replace(/"/g,'').trim());
            
            var val43 = parseFloat(calcCols[csvIndexIsExposed(25)] || 0);
            var val41 = parseFloat(calcCols[csvIndexIsExposed(20)] || 0);
            var diff = val43 - val41;
            diff = Math.abs(diff)
            calcCols[csvIndexIsExposed(25)] = diff === 0 ? "" : '"' + diff + '"';
            
            
            
            let good = 0;
            let bad = 0;
            let hold = 0;
            let inspect = 0;
            let label = 0;
            let total = 0;
            let avail = 0;
            let col14Val = calcCols[csvIndexIsExposed(18)];
            let col12Val = parseFloat(calcCols[csvIndexIsExposed(12)] || 0);
            let col19Val = diff;
            
            if (balances[itemid]){
              good  = parseFloat(balances[itemid].good);
              bad   = balances[itemid].bad;
              hold   = balances[itemid].hold;
              inspect   = balances[itemid].inspect;
              label   = balances[itemid].label;
              total = parseFloat(balances[itemid].total);
              avail = parseFloat(balances[itemid].avail);
            }
            
            
            if (itemid == 5472) {
              log.debug('good', good)
              log.debug('avail', avail)
              log.debug('total', total)
              log.debug('col12Val', col12Val)
              log.debug('balances', balances[itemid])              
              log.debug('Test', good - col12Val)   
            }
            
            var availtoProm = good - col12Val;
            
            
            calcCols[calcCols.length] = 'Black';
            calcCols[csvIndexIsExposed(12)] = '"' + availtoProm +'"';
            calcCols[csvIndexIsExposed(13)] = '"' + good + '"';
            calcCols[csvIndexIsExposed(14)] = '"' + bad + '"';
            calcCols[csvIndexIsExposed(15)] = '"' + inspect + '"';
            calcCols[csvIndexIsExposed(16)] = '"' + label + '"';
            calcCols[csvIndexIsExposed(17)] = '"' + hold + '"';
            calcCols[csvIndexIsExposed(18)] = '"' + total + '"';
            calcCols[csvIndexIsExposed(19)] = '"' + ((parseFloat(total)) / monthAvg).toFixed(2)+ '"';
            calcCols[csvIndexIsExposed(24)] = '"' + (((parseFloat(total)) + parseFloat(val41 || 0))/ monthAvg).toFixed(2)+'"';
            calcCols[csvIndexIsExposed(23)] = '"' + ((parseFloat(total)) + parseFloat(val41 || 0)).toFixed(2)+'"';
            calcCols[csvIndexIsExposed(27)] = '"' + ((parseFloat(total) + parseFloat(val43 || 0))/ monthAvg).toFixed(2)+'"';
            calcCols[csvIndexIsExposed(26)] = '"' + (parseFloat(total) + parseFloat(val43 || 0)).toFixed(2)+'"';
            calcCols[csvIndexIsExposed(31)] = '"' + avail + '"';
            function normalizeMovement(val) {
              // Null/undefined/empty string?
              if (val === null || val === undefined) return "No Movement";
              var s = String(val).replace(/,/g, "").trim();
              if (s === "") return "No Movement";
              
              // Parse number
              var n = parseFloat(s);
              if (!isFinite(n)) return "No Movement";
              
              // 0 or negative -> "0"; positive -> keep (or format)
              if (n <= 0) return "No Movement";
              return n.toFixed(2);
            }
            
            var col9 = normalizeMovement(monthAvg);
            const qtytotal = diff + safeParseFloat(calcCols[csvIndexIsExposed(20)]) + parseFloat(avail) - parseFloat(col12Val);
            const stockingQty = Math.ceil(parseFloat(calcCols[csvIndexIsExposed(11)]) * 4.5);
            calcCols[csvIndexIsExposed(11)] = '"' + col9 + '"';
            
            calcCols[csvIndexIsExposed(63)] = '"' + calcCols[csvIndexIsExposed(11)] + '"';
            calcCols[csvIndexIsExposed(64)] = '"' + parseFloat(calcCols[csvIndexIsExposed(11)]) * 4 + '"';
            
            let recommendedQty = 0;
            if (qtytotal < stockingQty) {
              recommendedQty = (stockingQty - qtytotal).toFixed(2);
            }
            
            
            
            const monthsStock = (diff + parseFloat(avail) + safeParseFloat(calcCols[csvIndexIsExposed(20)])) / monthAvg;
            
            if (calcCols[1] == 0) calcCols[1] = '""';
            calcCols[1] = '"' + recommendedQty + '"';
            
            let rowStyle = '';
            if (!isNaN(monthsStock) && monthsStock <= 4.5) {
              rowStyle = 'color: red;';
              calcCols[calcCols.length - 1] = 'Red';
            }
            
            // --- calcCols already built above ---
            // Pull the status once so we don't duplicate it
            var statusVal = (calcCols.length ? calcCols[calcCols.length - 1] : '') || '';
            // Remove the last element (status) from the working array
            var baseCols = calcCols.slice(0, -1);
            
            let displayCols;
            if (adminCsvIndex >= 0 && truncateAfterAdmin) {
              displayCols = baseCols.slice(0, adminCsvIndex);
            } else if (adminCsvIndex >= 0 && removeJustAdmin) {
              displayCols = baseCols.filter((_, i) => i !== adminCsvIndex);
            } else {
              displayCols = baseCols.slice();
            }
            displayCols.push(statusVal);
            
            
            
            function cleanCsvValue(value) {
              let v = value == null ? '' : String(value).trim();
              
              // 0) Normalize placeholders FIRST (before any quoting logic)
              
              // 1) If it's already wrapped in quotes, inspect the inner text
              if (v.length >= 2 && v[0] === '"' && v[v.length - 1] === '"') {
                const inner = v.slice(1, -1);
                
                // If inner contains special CSV chars or escaped quotes, keep it quoted (normalize escapes)
                if (/[",\r\n]/.test(inner) || /""/.test(inner)) {
                  return '"' + inner.replace(/"/g, '""') + '"';
                }
                
                // Otherwise it's safe to drop the outer quotes
                v = inner;
              }
              
              // 2) If there are stray unmatched quotes (e.g., starts with " but doesn't end with "),
              // strip them so we don't end up with duplicated quotes later.
              if ((v.startsWith('"') && !v.endsWith('"')) || (!v.startsWith('"') && v.endsWith('"'))) {
                v = v.replace(/^"+|"+$/g, '');
              }
              if ((v.startsWith('"') && !v.endsWith('"')) || (!v.startsWith('"') && v.endsWith('"'))) {
                v = v.replace(/^"+|"+$/g, '');
              }
              
              
              // 3) Quote only if needed (comma, quote, or newline present)
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
            
            
            // --- Render row (using displayCols):
            let rowId = newContent.length - 1; // keep same row id mapping as you had
            let rowHtml = '<tr style="' + rowStyle + '">';
            rowHtml += '<td class="col-sticky-1"><input type="checkbox" name="row_select_' + rowId + '" /></td>';
            rowHtml += '<td class="col-sticky-2"><input type="number" name="qty_input_' + rowId + '" min="0" value="' + recommendedQty + '" /></td>';
            rowHtml += '<td></td>';
            rowHtml += '<td></td>';
            rowHtml += '<td></td>';
            
            var first = true;
            displayCols.forEach(function (value) {
              if (first) {
                value = String(value || '').replace(/"/g, '').trim();
                if (value === '- None -' || value === 'NaN' || value === 'Infinity') value = '';
                rowHtml += '<td><input type="text" name="memo_input_' + rowId + '" value="' + value  + '" /></td>';
              }
              else {
                value = String(value || '').replace(/"/g, '').trim();
                if (value === '- None -' || value === 'NaN' || value === 'Infinity') value = '';
                if (value === '.00') value = '0.00';
                rowHtml += '<td>' + value + '</td>';
              }
              first = false;
              
            });
            
            rowHtml += '</tr>';
            
            var itemIdNum = parseInt(itemid, 10) || 0;
            
            if (rowStyle.includes('red')) {
              redRows.push({
                itemId: itemIdNum,
                html: rowHtml
              });
            } else {
              blackRows.push({
                itemId: itemIdNum,
                html: rowHtml
              });
            }
            
            
          });
          
          
          
          // Append red rows first, then black rows
          redRows.sort(function (a, b) {
            return a.itemId - b.itemId;
          });
          blackRows.sort(function (a, b) {
            return a.itemId - b.itemId;
          });
          
          // Append red rows first (sorted), then black rows (sorted)
          html += redRows.map(function (r) { return r.html; }).join('')
          + blackRows.map(function (r) { return r.html; }).join('');
          
          html += `</tbody></table></div></div></div>
          <div id="filter-portal"></div>
          
          <script>
          document.addEventListener('DOMContentLoaded', function () {
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
            
            brandSet.forEach(function (brand) {
              var opt = document.createElement('option');
              opt.value = brand;
              opt.textContent = brand;
              brandSelect.appendChild(opt);
            });
            
            brandCatSet.forEach(function (cat) {
              var opt = document.createElement('option');
              opt.value = cat;
              opt.textContent = cat;
              brandCatSelect.appendChild(opt);
            });
            
            deptSet.forEach(function (dept) {
              var opt = document.createElement('option');
              opt.value = dept;
              opt.textContent = dept;
              deptSelect.appendChild(opt);
            });
            
            polSet.forEach(function (pol) {
              var opt = document.createElement('option');
              opt.value = pol;
              opt.textContent = pol;
              polSelect.appendChild(opt);
            });
            
            productSet.forEach(function (product) {
              var opt = document.createElement('option');
              opt.value = product;
              opt.textContent = product;
              productSelect.appendChild(opt);
            });            
            
            itemSet.forEach(function (item) {
              var opt = document.createElement('option');
              opt.value = item;
              opt.textContent = item;
              itemSelect.appendChild(opt);
            });
            
            vendorSet.forEach(function (vendor) {
              var opt = document.createElement('option');
              opt.value = vendor;
              opt.textContent = vendor;
              vendorSelect.appendChild(opt);
            });
            
            function getSelectedValues(selectEl) {
              var selected = [];
              for (var i = 0; i < selectEl.options.length; i++) {
                if (selectEl.options[i].selected) {
                  selected.push(selectEl.options[i].value.toLowerCase());
                }
              }
              return selected;
            }
            function __vizIdxFromCsv(csvIdx){
              var controlOffset = 5; // 5 control columns before CSV data
              var hasAdmin = (typeof window.__ADMIN_IDX__ === 'number' && window.__ADMIN_IDX__ >= 0);
              
              // If we truncated after Admin (type=4), any csv index >= Admin is not visible at all.
              if (window.__TRUNC_AFTER__ && hasAdmin) {
                if (csvIdx >= window.__ADMIN_IDX__) return -1;         // not present in UI
                return controlOffset + csvIdx;                         // no shift needed
              }
              
              // If we just removed the Admin column (type=5), shift indexes >= Admin left by 1
              if (window.__ADMIN_REMOVED__ && hasAdmin) {
                var shift = (csvIdx >= window.__ADMIN_IDX__) ? 1 : 0;
                return controlOffset + csvIdx - shift;
              }
              
              // No change
              return controlOffset + csvIdx;
            }
            
            function getCellSafe(cells, idx){
              if (idx < 0) return '';                  // column not visible (e.g., truncated)
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
              
              document.querySelectorAll('#excelTable tbody tr').forEach(function (row) {
                var cells = row.querySelectorAll('td');
                var item    = getCellSafe(cells, __vizIdxFromCsv(${ITEM_INDEX}));
                var vendor  = getCellSafe(cells, __vizIdxFromCsv(${VENDOR_INDEX + 1}));
                var brand   = getCellSafe(cells, __vizIdxFromCsv(${BRAND_INDEX + 1}));
                var brandCat= getCellSafe(cells, __vizIdxFromCsv(${BRAND_CATEGORY_INDEX + 1}));
                var dept    = getCellSafe(cells, __vizIdxFromCsv(${DEPT_INDEX + 1}));
                var pol     = getCellSafe(cells, __vizIdxFromCsv(${POL_INDEX + 1}));
                var product     = getCellSafe(cells, __vizIdxFromCsv(${PRODUCT_INDEX + 1}));
                
                
                var itemMatch = selectedItems.length === 0 || selectedItems.includes(item);
                var vendorMatch = selectedVendors.length === 0 || selectedVendors.includes(vendor);
                var brandMatch = selectedBrands.length === 0 || selectedBrands.includes(brand);
                var brandCatMatch = selectedBrandCats.length === 0 || selectedBrandCats.includes(brandCat);
                var deptMatch = selectedDepts.length === 0 || selectedDepts.includes(dept);
                var polMatch = selectedPOL.length === 0 || selectedPOL.includes(pol);
                var productMatch = selectedPRODUCT.length === 0 || selectedPRODUCT.includes(product);
                
                row.style.display = itemMatch && vendorMatch && brandMatch && brandCatMatch && deptMatch &&  polMatch &&  productMatch ? '' : 'none';
              });
            }
            
            itemSelect.addEventListener('change', filterTable);
            vendorSelect.addEventListener('change', filterTable);
            brandSelect.addEventListener('change', filterTable);
            brandCatSelect.addEventListener('change', filterTable);
            deptSelect.addEventListener('change', filterTable);
            polSelect.addEventListener('change', filterTable);
            productSelect.addEventListener('change', filterTable);
          });
          </script>
          <script>
          document.addEventListener('DOMContentLoaded', function () {
            // Make the export button open the cleaned CSV in a new tab
            var exportBtn = document.getElementById('downloadCsvBtn');
            if (exportBtn) {
              exportBtn.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();
                var url = (window.DOWNLOAD_URL || '').trim(); // value will be set later
                if (url) {
                  window.open(url, '_blank');
                } else {
                  alert('Download not available yet.');
                }
              });
            }
            
            
          });
          </script>
          <script>
          document.addEventListener('DOMContentLoaded', function () {
            var suiteletForm = document.querySelector('form');
            var hiddenSelected = document.getElementById('custpage_selected'); // NetSuite renders the field with this id
            
            function collectSelectedRows() {
              if (!hiddenSelected) return;
              var out = [];
              var rows = document.querySelectorAll('#excelTable tbody tr');
              rows.forEach(function (tr) {
                var cb = tr.querySelector('input[type="checkbox"][name^="row_select_"]');
                if (cb && cb.checked) {
                  var rowId = (cb.name || '').split('_').pop();
                  var qtyInput = tr.querySelector('input[type="number"][name^="qty_input_"]');
                  var qty = qtyInput ? (qtyInput.value || '0') : '0';
                  
                  var memoInput = tr.querySelector('input[type="text"][name^="memo_input_"]');
                  var memo = memoInput ? (memoInput.value || '') : '';
                  
                  
                  var mosVal = '';
                  var mosCell = tr.cells[2];
                  if (mosCell) {
                    mosVal = (mosCell.textContent || '').trim();
                  }
                  
                  out.push({ rowId: rowId, qty: qty, memo: memo, mos: mosVal });
                  
                }
              });
              hiddenSelected.value = JSON.stringify(out);
            }
            
            function closeHeaderPanel() {
              var open = document.querySelector('#filter-portal .th-filter-panel');
              if (open) open.remove();
              document.querySelectorAll('th.th-filter-active').forEach(function (th) {
                th.classList.remove('th-filter-active');
              });
            }
            
            function nativeSubmit() {
              if (!suiteletForm) return;
              // 1) capture selections
              collectSelectedRows();
              // 2) clean overlays
              closeHeaderPanel();
              // 3) force native submit
              try {
                HTMLFormElement.prototype.submit.call(suiteletForm);
              } catch (e) {
                if (typeof suiteletForm.submit === 'function') suiteletForm.submit();
              }
            }
            
            var nsSubmitBtn = document.querySelector('input[type="submit"], button[type="submit"]');
            if (nsSubmitBtn) {
              nsSubmitBtn.addEventListener('click', function () {
                // queue native submit to run after other handlers
                setTimeout(nativeSubmit, 0);
              }, true);
            }
            
            if (suiteletForm) {
              suiteletForm.addEventListener('submit', function () {
                // safety: if something intercepts submit, still collect & force
                setTimeout(nativeSubmit, 0);
              }, true);
            }
          });
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
          
          var reloadFile = file.load({id: newFileId2});
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
          + 'window.__ADMIN_IDX__='     + JSON.stringify(adminCsvIndex) + ';'
          + 'window.__TRUNC_AFTER__='   + JSON.stringify(truncateAfterAdmin && adminCsvIndex >= 0) + ';'
          + 'window.__ADMIN_REMOVED__=' + JSON.stringify(removeJustAdmin    && adminCsvIndex >= 0) + ';'
          + '</script>';
          
          
          
          
          html += `
          <script>
          document.addEventListener('DOMContentLoaded', function () {
            // --- DOM refs
            var container = document.querySelector('.table-container');
            var table = document.getElementById('excelTable');
            var topScroll = document.getElementById('topScroll');
            var topInner = document.getElementById('topScrollInner');
            if (!table) return;
            
            // ===== PERF FLAGS (NEW) =====
            var __widthsApplied = false;     // heavy: sets widths on many cells
            var __stickyClassesApplied = false; // heavy: adds classes to many cells
            
            
            // --- Fixed widths (px) for the first N columns (sticky/frozen)
            // Adjust these numbers to your exact needs:
            var FIXED_STICKY_WIDTHS = [
              50,   // 0: Select
              100,  // 1: Order Qty
              110,  // 2: Month of Stock
              110,  // 3: Item Cubic Space
              90,   // 4: Weight
              200,  // 5: CSV Col 1 (first data col)
              110,  // 6: CSV Col 2
              180,  // 7: CSV Col 3
              90,   // 8: CSV Col 4
              450   // 9: CSV Col 5
            ]; // <-- first 9 frozen columns total
            
            // Apply explicit width/min/max-width to first N columns
            function applyFixedWidths(widths){
              if (!table || !table.tHead || !table.tBodies[0]) return;
              
              // Always keep header widths correct (cheap)
              var ths = table.tHead.rows[0].cells;
              for (var i = 0; i < widths.length && i < ths.length; i++){
                var w = widths[i];
                ths[i].style.width = w + 'px';
                ths[i].style.minWidth = w + 'px';
                ths[i].style.maxWidth = w + 'px';
              }
              
              // Heavy part: only once
              if (__widthsApplied) return;
              __widthsApplied = true;
              
              var rows = table.tBodies[0].rows;
              for (var r = 0; r < rows.length; r++){
                var tds = rows[r].cells;
                for (var c = 0; c < widths.length && c < tds.length; c++){
                  var w2 = widths[c];
                  tds[c].style.width = w2 + 'px';
                  tds[c].style.minWidth = w2 + 'px';
                  tds[c].style.maxWidth = w2 + 'px';
                  tds[c].style.overflow = 'hidden';
                  tds[c].style.textOverflow = 'ellipsis';
                }
              }
            }
            
            
            // ========== Top scrollbar ==========
            function updateTopScrollbarWidth(){
              if (!container || !table || !topInner) return;
              var w = Math.max((table.scrollWidth || 0), (container.clientWidth || 0) + 2);
              topInner.style.width = w + 'px';
            }
            if (topScroll) topScroll.style.display = 'block';
            if (topScroll && container) {
              topScroll.addEventListener('scroll', function(){ container.scrollLeft = topScroll.scrollLeft; });
              container.addEventListener('scroll', function(){ topScroll.scrollLeft = container.scrollLeft; });
            }
            
            // ========== Sticky N columns ==========
            function makeSticky(firstN){
              if (!table || !table.tHead || !table.tHead.rows.length) return;
              
              // Build left offsets from FIXED_STICKY_WIDTHS (fallback to measured if missing)
              var lefts = [];
              var acc = 0;
              
              for (var i = 0; i < firstN; i++) {
                var w = FIXED_STICKY_WIDTHS[i];
                if (typeof w !== 'number' || isNaN(w)) {
                  // fallback: measure header if width not provided
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
                  
                  // Heavy classList.add only once
                  if (!__stickyClassesApplied) {
                    cell.classList.add('sticky-col');
                    if (c > 0) cell.classList.add('sep-left');
                  }
                  
                  // Always update left + zIndex (needed)
                  cell.style.left = lefts[c] + 'px';
                  cell.style.zIndex = (r === 0 ? 30 : 20) + c;
                }
              }
              __stickyClassesApplied = true;
              
            }
            
            
            // ========== Excel-style header filters ==========
            var activeFilters = Object.create(null); // { colIndex: Set(lowercase values) }
            
            function bodyRows(){ return table.tBodies && table.tBodies[0] ? table.tBodies[0].rows : []; }
            function getCellText(row, idx){
              var cells = row.cells; if (!cells || idx < 0 || idx >= cells.length) return '';
              var t = cells[idx].textContent || ''; return String(t).trim();
            }
            function normalizeVal(v){
              // Treat null/undefined/"- None -"/"NaN" as blank
              var s = String(v == null ? '' : v).trim();
              if (/^-+\s*none\s*-+$/i.test(s)) s = '';
              if (/^nan$/i.test(s)) s = '';
              return s.toLowerCase();
            }
            
            // De-duped display values in the column (each value appears once)
            function getAllValuesForColumn(colIdx){
              var rows = bodyRows();
              var displayByKey = Object.create(null);
              
              for (var r=0; r<rows.length; r++){
                var v = getCellText(rows[r], colIdx);
                // Normalize AND clean per your existing blank rules
                if (v === '- None -' || v === 'NaN') v = '';
                if (v === '.00') v = '0.00';
                
                var key = normalizeVal(v);
                // First-seen display value wins; blanks get "(blank)"
                if (!displayByKey[key]) displayByKey[key] = (key === '' ? '(blank)' : (v || '(blank)'));
              }
              
              // Sort keys by display label (blank at bottom)
              var keys = Object.keys(displayByKey).sort(function(a,b){
                var da = displayByKey[a], db = displayByKey[b];
                if (a==='' && b==='') return 0;
                if (a==='') return 1;     // put blank last
                if (b==='') return -1;
                return da.localeCompare(db, undefined, {numeric:true, sensitivity:'base'});
              });
              
              return { keys: keys, displayByKey: displayByKey };
            }
            
            // Row passes all filters EXCEPT the specified column
            function rowPassesFiltersExceptColumn(row, exceptColIdx){
              for (var k in activeFilters){
                if (!Object.prototype.hasOwnProperty.call(activeFilters,k)) continue;
                var idx = parseInt(k,10);
                if (idx === exceptColIdx) continue;
                var set = activeFilters[k];
                if (set && set.size > 0){
                  var v = getCellText(row, idx).toLowerCase();
                  if (!set.has(v)) return false;
                }
              }
              return true;
            }
            
            // Counts per value for a column among rows that pass other filters
            function getValueCountsForColumn(colIdx){
              var rows = bodyRows(), counts = Object.create(null);
              for (var r=0; r<rows.length; r++){
                var row = rows[r];
                if (!rowPassesFiltersExceptColumn(row, colIdx)) continue;
                
                var v = getCellText(row, colIdx);
                if (v === '- None -' || v === 'NaN') v = '';
                if (v === '.00') v = '0.00';
                
                var key = normalizeVal(v);
                counts[key] = (counts[key] || 0) + 1;
              }
              return counts; // { normKey: count }
            }
            // Apply all active filters
            function applyFilters(){
              var rows = bodyRows(), pairs = [];
              for (var k in activeFilters){
                if (!Object.prototype.hasOwnProperty.call(activeFilters,k)) continue;
                var s = activeFilters[k];
                if (s && s.size > 0) pairs.push([parseInt(k,10), s]);
              }
              for (var r=0; r<rows.length; r++){
                var row = rows[r], show = true;
                for (var i=0; i<pairs.length && show; i++){
                  var colIdx = pairs[i][0], set = pairs[i][1];
                  var val = getCellText(row, colIdx).toLowerCase();
                  if (!set.has(val)) show = false;
                }
                row.style.display = show ? '' : 'none';
              }
            }
            
            // Close panels when clicking outside
            document.addEventListener('click', function(e){
              var open = table.querySelectorAll('.th-filter-panel');
              for (var i=0;i<open.length;i++){
                var th = open[i].closest('th');
                var trigger = th && th.querySelector('.th-filter-btn');
                var inside = open[i].contains(e.target) || (trigger && trigger.contains(e.target));
                if (!inside){ if (th) th.classList.remove('th-filter-active'); open[i].remove(); }
              }
            });
            
            function openFilterPanel(btn, colIdx){
              var th = btn.closest('th');
              var existing = document.querySelector('#filter-portal .th-filter-panel');
              if (existing){
                // if panel is already for the same TH, just close; otherwise close and continue
                var sameTh = existing.__ownerTH === th;
                existing.remove();
                th && th.classList.remove('th-filter-active');
                if (sameTh) return;
              }
              
              // Close any panel-activated state on other THs
              table.querySelectorAll('th.th-filter-active').forEach(function(node){
                node.classList.remove('th-filter-active');
              });
              
              // --- Build panel
              var panel = document.createElement('div');
              panel.className = 'th-filter-panel';
              panel.innerHTML =
              '<div class="hdr">Filter</div>' +
              '<input type="search" placeholder="Search values...">' +
              '<div class="list"></div>' +
              '<div class="actions">' +
              '<button type="button" data-act="clear">Clear</button>' +
              '<div style="display:flex; gap:6px;">' +
              '<button type="button" data-act="selectall">Select all</button>' +
              '<button type="button" data-act="deselectall">Deselect all</button>' +
              '<button type="button" data-act="apply">Apply</button>' +
              '</div>' +
              '</div>';
              var list = panel.querySelector('.list');
              
              // Keep a back-reference so we know which TH owns this panel
              panel.__ownerTH = th;
              
              // --- Compute data for this column (your existing helpers are reused)
              // Build data for this column (grouped)
              var grouped = getAllValuesForColumn(colIdx);     // { keys, displayByKey }
              var universeKeys = grouped.keys;                 // array of normalized keys
              var displayByKey = grouped.displayByKey;        // map key -> label to show once
              var visibleCounts = getValueCountsForColumn(colIdx); // counts by key
              var existingAllowed = activeFilters[colIdx] || null;
              
              // Temp selection stores *keys* (normalized)
              var tempSelection = new Set();
              
              // Default selection: values visible given other filters, or previously applied
              if (existingAllowed && existingAllowed.size > 0){
                universeKeys.forEach(function(k){
                  if (existingAllowed.has(k)) tempSelection.add(k);
                });
              } else {
                universeKeys.forEach(function(k){
                  if ((visibleCounts[k] || 0) > 0) tempSelection.add(k);
                });
              }
              
              // Render list: one row per *key*, using displayByKey[key]
              function renderList(filterText){
                list.innerHTML = '';
                var ft = String(filterText || '').toLowerCase();
                
                universeKeys.forEach(function(k){
                  var labelText = displayByKey[k] || (k === '' ? '(blank)' : k);
                  if (ft && labelText.toLowerCase().indexOf(ft) === -1) return;
                  
                  var id = 'f_' + colIdx + '_' + Math.random().toString(36).slice(2);
                  var checked = tempSelection.has(k);
                  var cnt = visibleCounts[k] || 0;
                  var row = document.createElement('div'); row.className = 'row';
                  row.dataset.key = k; // store normalized key
                  row.innerHTML =
                  '<input type="checkbox" id="'+id+'" '+(checked?'checked':'')+'>' +
                  '<label for="'+id+'">'+ labelText + (cnt ? '  ('+cnt+')' : '') +'</label>';
                  list.appendChild(row);
                });
              }
              renderList('');
              
              // Persist selection while typing/toggling
              panel.querySelector('input[type="search"]').addEventListener('input', function(){
                renderList(this.value);
              });
              list.addEventListener('change', function(e){
                var cb = e.target; if (!cb || cb.type !== 'checkbox') return;
                var row = cb.closest('.row'); if (!row) return;
                var key = row.dataset.key || '';
                if (cb.checked) tempSelection.add(key); else tempSelection.delete(key);
              });
              
              // Buttons
              panel.querySelector('.actions').addEventListener('click', function(e){
                var act = e.target && e.target.getAttribute('data-act'); if (!act) return;
                
                if (act === 'clear'){
                  delete activeFilters[colIdx];
                  th.classList.remove('th-filter-active');
                  panel.remove();
                  applyFilters();
                  return;
                }
                
                if (act === 'selectall'){
                  var cbs = list.querySelectorAll('input[type="checkbox"]');
                  for (var i=0;i<cbs.length;i++){
                    var cb = cbs[i];
                    if (!cb.checked) cb.checked = true;
                    var key = cb.closest('.row').dataset.key || '';
                    tempSelection.add(key);
                  }
                  return;
                }
                
                if (act === 'deselectall'){                     // <— add this block
                  var cbs = list.querySelectorAll('input[type="checkbox"]');
                  for (var i=0;i<cbs.length;i++){
                    var cb = cbs[i];
                    if (cb.checked) cb.checked = false;
                    var key = cb.closest('.row').dataset.key || '';
                    tempSelection.delete(key);
                  }
                  return;
                }
                
                if (act === 'apply'){
                  var allow = new Set();
                  tempSelection.forEach(function(k){ allow.add(k); });
                  
                  // Selecting *all keys* is equivalent to "no filter"
                  if (allow.size === universeKeys.length){
                    delete activeFilters[colIdx];
                    th.classList.remove('th-filter-active');
                  } else {
                    activeFilters[colIdx] = allow;
                    th.classList.add('th-filter-active');
                  }
                  panel.remove();
                  applyFilters();
                }
              });
              
              
              // --- Mount into portal (outside of table to avoid clipping)
              var portal = document.getElementById('filter-portal') || document.body;
              portal.appendChild(panel);
              
              // --- Position the panel relative to the button (above all fixed/sticky cells)
              function positionPanel(){
                var rect = btn.getBoundingClientRect();
                var vw = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
                var ph = panel.offsetHeight || 0;
                var pw = panel.offsetWidth  || 260;
                
                // Prefer dropping below header; if not enough space, show above
                var belowTop = rect.bottom + 6;
                var aboveTop = rect.top - ph - 6;
                var top = (belowTop + ph <= window.innerHeight) ? belowTop : Math.max(8, aboveTop);
                
                // Right-align to button’s right edge, but keep on-screen
                var left = Math.min(vw - pw - 8, Math.max(8, rect.right - pw));
                
                panel.style.top  = Math.round(top)  + 'px';
                panel.style.left = Math.round(left) + 'px';
              }
              positionPanel();
              
              // Reposition while scrolling / resizing / table scrolling
              
              var __posRaf = false;
              function schedulePosition(){
                if (__posRaf) return;
                __posRaf = true;
                requestAnimationFrame(function(){
                  __posRaf = false;
                  positionPanel();
                });
              }
              var onScrollOrResize = schedulePosition;
              
              window.addEventListener('scroll', onScrollOrResize, true);
              window.addEventListener('resize', onScrollOrResize);
              if (container) container.addEventListener('scroll', onScrollOrResize);
              
              
              // Clean up listeners when panel closes
              function teardown(){
                window.removeEventListener('scroll', onScrollOrResize, true);
                window.removeEventListener('resize', onScrollOrResize);
                if (container) container.removeEventListener('scroll', onScrollOrResize);
              }
              // Close on outside click
              setTimeout(function(){
                document.addEventListener('click', function handler(e){
                  if (panel.contains(e.target) || th.contains(e.target)) return;
                  panel.remove(); th.classList.remove('th-filter-active'); teardown();
                  document.removeEventListener('click', handler);
                });
              }, 0);
              
              th.classList.add('th-filter-active');
            }
            
            
            // Delegate header clicks
            if (table.tHead){
              table.tHead.addEventListener('click', function(e){
                var btn = e.target.closest('.th-filter-btn');
                if (!btn) return;
                var colIdx = parseInt(btn.getAttribute('data-col'), 10);
                if (!isFinite(colIdx)) return;
                e.stopPropagation();
                openFilterPanel(btn, colIdx);
              });
            }
            
            // ========== Init ==========
            function init(){
              updateTopScrollbarWidth();
              
              // 1) Lock widths for the first 9 frozen columns
              applyFixedWidths(FIXED_STICKY_WIDTHS);
              
              // 2) Then compute sticky lefts using those fixed widths
              makeSticky(11); // 5 control + first 5 data cols
              
              if (topScroll && container) topScroll.scrollLeft = container.scrollLeft;
            }
            requestAnimationFrame(init);
            setTimeout(init, 200);
            
            // Debounced resize (NEW)
            var __resizeT = null;
            window.addEventListener('resize', function(){
              clearTimeout(__resizeT);
              __resizeT = setTimeout(init, 180);
            });
            
          });
          </script>
          
          `;
          
          htmlField.defaultValue = html;
          form.clientScriptModulePath = './CL_RR_Tool.js'; // Update path as needed
          
          // form.addButton({
          //     id: 'custpage_download_btn',
          //     label: 'Download Cleaned File',
          //     functionName: 'downloadFile('+ newFileId +')'
          // });
          context.response.writePage(form);
        }
        
        if (context.request.method === 'POST') {
          var params = context.request.parameters;
          var fileId = params.custpage_file_id;
          
          var p = context.request.parameters || {};
          var postedEmp = p.custpage_empid || '';
          var postedTs = p.custpage_ts || '';
          var postedSig = p.custpage_sig || '';
          
          // Allow legacy (no signature) if you want; otherwise require verify(postedEmp, postedTs, postedSig)
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
                var picked = JSON.parse(selectedJson); // [{rowId, qty}]
                picked.forEach(function (p) {
                  var rowId = String(p.rowId || '').trim();
                  var qty = parseInt(p.qty, 10) || 0;
                  var memo = p.memo || 0;
                  var mos   = parseFloat(p.mos);
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
                  log.debug('params', params)
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
              log.debug('cols', cols)
              log.debug('cols', cols)
              
              try {
                record.create({
                  type: 'customrecord_mi_planned_po',
                  isDynamic: true
                })
                .setValue({ fieldId: 'custrecord_mi_item', value: cols[3] })
                .setValue({ fieldId: 'custrecord_mi_order_qty', value: entry.qty })
                .setValue({ fieldId: 'custrecord_mi_purchase_memo', value: entry.memo == 0 ? '' : entry.memo})
                .setValue({ fieldId: 'custrecord_month_of_stocks', value: safeParseFloat(entry.monthStock) }) // <-- NEW
                .setValue({ fieldId: 'custrecord_mi_qty_of_ordered_not_ship', value: safeParseFloat(cols[24]) })
                .setValue({ fieldId: 'custrecord_mi_qty_available', value: safeParseFloat(cols[31]) })
                .setValue({ fieldId: 'custrecord_mi_qty_in_transit', value: safeParseFloat(cols[25]) })
                .setValue({ fieldId: 'custrecord_mi_min_month_qty', value: safeParseFloat(cols[16]) })
                .save();
                
                createdCount++; // Increment on success
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
                font-family: Arial, sans-serif;
                text-align: center;
                padding-top: 100px;
              }
              .message-box {
                display: inline-block;
                background-color: #f3f9ff;
                padding: 20px 30px;
                border: 1px solid #a3d2f2;
                border-radius: 8px;
                color: #005b99;
                font-size: 16px;
              }
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
                ["status","anyof","6","1","3","5","8"], 
                // "AND", 
                // [["onhand","greaterthan","0"],"OR",["available","greaterthan","0"]], 
                "AND", 
                ["item.isinactive", "is", "F"]
              ],
              columns: [
                search.createColumn({
                  name: "item",
                  summary: "GROUP"
                }),
                search.createColumn({
                  name: "formulanumeric",
                  summary: "SUM",
                  formula: "case when {status} = 'Good' then {onhand} else 0 end"
                }),
                search.createColumn({
                  name: "formulanumeric1",
                  summary: "SUM",
                  formula: "case when {status} = 'Deviation' then {onhand} else 0 end"
                }),
                search.createColumn({
                  name: "formulanumeric2",
                  summary: "SUM",
                  formula: "case when {status} = 'Hold' then {onhand} else 0 end"
                }),
                search.createColumn({
                  name: "formulanumeric3",
                  summary: "SUM",
                  formula: "case when {status} = 'Inspection' then {onhand} else 0 end"
                }),
                search.createColumn({
                  name: "formulanumeric4",
                  summary: "SUM",
                  formula: "case when {status} = 'Label' then {onhand} else 0 end"
                }),
                search.createColumn({
                  name: "available",
                  summary: "SUM",
                  label: "Available"
                })
              ]
            });
            
            inventorybalanceSearchObj.run().each(function(result) {
              var itemId = result.getValue({ name: "item", summary: "GROUP" });
              var goodQty = parseFloat(result.getValue({ name: "formulanumeric", summary: "SUM" })) || 0;
              var badQty  = parseFloat(result.getValue({ name: "formulanumeric1", summary: "SUM" })) || 0;
              var holdQty  = parseFloat(result.getValue({ name: "formulanumeric2", summary: "SUM" })) || 0;
              var inspectQty  = parseFloat(result.getValue({ name: "formulanumeric3", summary: "SUM" })) || 0;
              var labelQty  = parseFloat(result.getValue({ name: "formulanumeric4", summary: "SUM" })) || 0;
              var total = parseFloat((goodQty + badQty).toFixed(2));
              var avail = parseFloat(result.getValue({ name: "available", summary: "SUM" })) || 0;
              
              resultMap[itemId] = { good: goodQty, bad: badQty, hold: holdQty, inspect: inspectQty, label: labelQty, total: total, avail: avail };
              return true;
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