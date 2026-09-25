/**
* @NApiVersion 2.1
* @NScriptType Suitelet
*/
define(['N/ui/serverWidget', 'N/search', 'N/log', 'N/runtime','N/file','N/crypto'],
function (ui, search, log, runtime, file, crypto) {

  function onRequest(context) {
    if (context.request.method === 'GET') {

      // --- Signed session helpers ---
      const SECRET = runtime.getCurrentScript().getParameter({ name: 'custscript_portal_secret' }) || 'change-me';
      const TOKEN_TTL_MS = 30 * 60 * 1000; // 30 minutes
      function sign(empid, ts) {
        const h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
        h.update({ input: empid + '|' + ts + '|' + SECRET });
        return h.digest({ outputEncoding: crypto.Encoding.HEX });
      }
      function verify(empid, ts, sig) {
        if (!empid || !ts || !sig) return false;
        if (Math.abs(Date.now() - parseInt(ts, 10)) > TOKEN_TTL_MS) return false;
        try { return sign(empid, ts) === sig; } catch (e) { log.error('verify token', e); return false; }
      }
      function toIdArray(val) {
        if (!val) return [];
        if (Array.isArray(val)) return val.map(v => String(v.value || v)).filter(Boolean);
        return String(val).split(',').map(s => s.trim()).filter(Boolean);
      }

      var form = ui.createForm({ title: 'Item Inventory Dashboard' });

      var ClientID = runtime.getCurrentScript().getParameter({ name: 'custscript_mi_client_id' });
      log.debug('ClientID', ClientID);

      // -------- Read incoming params (legacy + signed) --------
      var q = context.request.parameters || {};
      var selectedItem = q.custpage_item;
      var selectedItemId = q.custpage_itemid;
      var selectedProductCategory = q.custpage_productcategory;
      var selectedBrand = q.custpage_brand;

      // Legacy params
      var selectedEmp;
      var pricelevel = q.custpage_price || '';     // legacy csv of price level ids

      // Signed params (preferred)
      var empid = q.empid || '';
      var ts = q.ts || '';
      var sig = q.sig || '';

      // If signature present and valid, trust it and derive employee + price level
      if (empid && ts && sig && verify(empid, ts, sig)) {
        selectedEmp = empid; // override legacy
        // Get price level from employee if not provided
        try {
          var lf = search.lookupFields({
            type: search.Type.EMPLOYEE,
            id: parseInt(selectedEmp, 10),
            columns: ['custentity_mi_price_level']
          });
          if (lf && lf.custentity_mi_price_level) {
            // lookupFields returns string or array; normalize to csv string for downstream
            var arr = toIdArray(lf.custentity_mi_price_level);
            pricelevel = arr.join(',');
          }
        } catch (e) {
          log.error('lookupFields price level error', e);
        }
      }

      log.debug('selectedEmp', selectedEmp);
      log.debug('params', q);

      // No auth → kick back to portal
      if (!selectedEmp) {
        context.response.write(`
          <html>
          <head>
          <script type="text/javascript">
          setTimeout(function() {
            window.location.href = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2110&deploy=1&compid=4975346&ns-at=AAEJ7tMQamzukv1WMqTK6i2c27bRetbrd2MDLjhDgPPFOawMxCo';
          }, 1500);
          </script>
          <style>
            body { display:flex; justify-content:center; align-items:center; height:100vh; font-family:Arial, sans-serif; background:#0b0b0b; color:#fff; }
            .message { font-size:22px; font-weight:700; }
          </style>
          </head>
          <body><div class="message">Login Required</div></body>
          </html>
        `);
        return;
      }

      // ---- Hidden employee field ----
      var empField = form.addField({
        id: 'custpage_empid',
        label: 'Current Employee',
        type: ui.FieldType.SELECT,
        source: 'employee'
      });
      empField.defaultValue = selectedEmp;
      empField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

      // ---- Normalize price level ids (from legacy csv or from lookup above) ----
      var pricelevelids = toIdArray(pricelevel);
      log.debug('pricelevel', pricelevel);
      log.debug('pricelevelids', pricelevelids);

      var sigEmpField = form.addField({ id: 'custpage_sig_empid', label: 'Sig EmpID', type: ui.FieldType.TEXT });
      sigEmpField.defaultValue = empid || selectedEmp || '';
      sigEmpField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

      var sigTsField = form.addField({ id: 'custpage_sig_ts', label: 'Sig Timestamp', type: ui.FieldType.TEXT });
      sigTsField.defaultValue = ts || String(Date.now());
      sigTsField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

      var sigHashField = form.addField({ id: 'custpage_sig_hmac', label: 'Sig HMAC', type: ui.FieldType.TEXT });
      sigHashField.defaultValue = sig || (selectedEmp ? sign(selectedEmp, sigTsField.defaultValue) : '');
      sigHashField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });


      // ---- Filters UI ----
      var itemField = form.addField({ id: 'custpage_item', label: 'Item Name', type: ui.FieldType.SELECT });
      itemField.addSelectOption({ value: '', text: '-- Select Item Name --' });

      var itemIdField = form.addField({ id: 'custpage_itemid', label: 'Item Internal ID', type: ui.FieldType.SELECT });
      itemIdField.addSelectOption({ value: '', text: '-- Select Item ID --' });

      var productCategoryField = form.addField({ id: 'custpage_productcategory', label: 'Product Category', type: ui.FieldType.SELECT });
      productCategoryField.addSelectOption({ value: '', text: '-- Select Product Category --' });

      var brandField = form.addField({ id: 'custpage_brand', label: 'Brand', type: ui.FieldType.SELECT });
      brandField.addSelectOption({ value: '', text: '-- Select Brand --' });

      // ---- Brand dropdown (respects price levels) ----
      var brandfilters = [
        ['isinactive', 'is', 'F'], 'AND',
        ['custitem_mi_include_on_inv_dash', 'is', 'T']
      ];
      if (pricelevelids.length > 0) {
        brandfilters.push('AND', ['pricing.pricelevel', 'anyof', pricelevelids]);
      }
      var brandSearch = search.create({
        type: search.Type.ITEM,
        filters: brandfilters,
        columns: [search.createColumn({ name: 'cseg_mi_brand', summary: 'GROUP' })]
      });
      brandSearch.run().each(function (res) {
        var brandValue = res.getValue({ name: 'cseg_mi_brand', summary: 'GROUP' });
        var brandText = res.getText({ name: 'cseg_mi_brand', summary: 'GROUP' });
        if (brandValue && brandText) brandField.addSelectOption({ value: brandValue, text: brandText });
        return true;
      });

      // ---- Item / ID dropdowns (respects price levels) ----
      var dropdownfilters = [
        ['isinactive', 'is', 'F'], 'AND',
        ['custitem_mi_include_on_inv_dash', 'is', 'T']
      ];
      if (pricelevelids.length > 0) {
        dropdownfilters.push('AND', ['pricing.pricelevel', 'anyof', pricelevelids]);
      }
      var dropdownSearch = search.create({
        type: search.Type.ITEM,
        filters: dropdownfilters,
        columns: [
          search.createColumn({ name: 'itemid', summary: 'GROUP' }),
          search.createColumn({ name: 'internalid', summary: 'GROUP' })
        ]
      });
      dropdownSearch.run().each(function (res) {
        var id = res.getValue({ name: 'internalid', summary: 'GROUP' });
        var name = res.getValue({ name: 'itemid', summary: 'GROUP' });
        if (name) {
          itemField.addSelectOption({ value: id, text: name });
          itemIdField.addSelectOption({ value: id, text: id });
        }
        return true;
      });

      // ---- Product category dropdown (respects price levels) ----
      var categoryfilters = [
        ['isinactive', 'is', 'F'], 'AND',
        ['custitem_mi_include_on_inv_dash', 'is', 'T']
      ];
      if (pricelevelids.length > 0) {
        categoryfilters.push('AND', ['pricing.pricelevel', 'anyof', pricelevelids]);
      }
      var categorySearch = search.create({
        type: search.Type.ITEM,
        filters: categoryfilters,
        columns: [search.createColumn({ name: 'custitem_mi_product_category', summary: 'GROUP' })]
      });
      categorySearch.run().each(function (res) {
        var catValue = res.getValue({ name: 'custitem_mi_product_category', summary: 'GROUP' });
        var catText = res.getText({ name: 'custitem_mi_product_category', summary: 'GROUP' });
        if (catValue && catText) productCategoryField.addSelectOption({ value: catValue, text: catText });
        return true;
      });

      form.addSubmitButton({ label: 'Export CSV' });

      // ---- Main list search ----
      var itemTypes = ['InvtPart', 'Assembly', 'NonInvtPart'];
      var filters = [['type', 'anyof', itemTypes],'AND', ['custitem_mi_include_on_inv_dash', 'is', 'T'], 'AND', [["isinactive","is","F"],"OR",[["custitem_mi_available_sale","is","T"],"OR",["custitem_molisana_status","anyof","3"]]]];

      if (selectedItem) { filters.push('AND', ['internalid', 'anyof', selectedItem]); itemField.defaultValue = selectedItem; }
      if (pricelevelids.length > 0) { filters.push('AND', ['pricing.pricelevel', 'anyof', pricelevelids]); }
      if (selectedItemId) { filters.push('AND', ['internalid', 'anyof', selectedItemId]); itemIdField.defaultValue = selectedItemId; }
      if (selectedProductCategory) { filters.push('AND', ['custitem_mi_product_category', 'anyof', selectedProductCategory]); productCategoryField.defaultValue = selectedProductCategory; }
      if (selectedBrand) { filters.push('AND', ['cseg_mi_brand', 'anyof', selectedBrand]); brandField.defaultValue = selectedBrand; }

      var itemSearch = search.create({
        type: search.Type.ITEM,
        filters: filters,
        columns: [
          search.createColumn({name: 'internalid', sort: search.Sort.ASC}),
          'itemid',
          'displayname',
          'type',
          'cseg_mi_brand',
          'custitem_mi_cr_itm_cat',
          'custitem_mi_product_category',
          'baseprice',
          search.createColumn({ name: 'pricelevel', join: 'pricing' }),
          search.createColumn({
            name: "formulanumeric",
            formula: "CASE WHEN {custitem_esclusive_price_level} IS NULL OR {custitem_esclusive_price_level} = '' THEN {pricing.unitprice} WHEN {custitem_esclusive_price_level} = {pricing.pricelevel} THEN {pricing.unitprice} END",
            label: "Formula (Numeric)"
          })
        ]
      });

      var sublist = form.addSublist({ id: 'custpage_itemlist', label: 'Item Inventory Details', type: ui.SublistType.LIST });

      // Run & group
      var itemSearchResults = {};
      var priceLevels = {};
      var pagedResults = itemSearch.runPaged({ pageSize: 2500 });
      pagedResults.pageRanges.forEach(function (pageRange) {
        var pageData = pagedResults.fetch({ index: pageRange.index });
        pageData.data.forEach(function (res) {
          var itemId = res.getValue('internalid');
          var priceLevelName = res.getText({ name: 'pricelevel', join: 'pricing' });
          var unitPrice = res.getValue({ name: 'formulanumeric' });

          if (!itemSearchResults[itemId]) {
            itemSearchResults[itemId] = {
              itemid: res.getValue('itemid'),
              internalid: itemId,
              displayname: res.getValue('displayname'),
              custitem_mi_cr_itm_cat: res.getText('custitem_mi_cr_itm_cat'),
              custitem_mi_product_category: res.getText('custitem_mi_product_category'),
              cseg_mi_brand: res.getText('cseg_mi_brand'),
              prices: {}
            };
          }
          if (priceLevelName) {
            itemSearchResults[itemId].prices[priceLevelName] = unitPrice;
            priceLevels[priceLevelName] = true;
          }
          return true;
        });
      });

      // Columns
      var sublistFieldMeta = {};
      [
        { id: 'itemid', label: 'Item Name', type: ui.FieldType.TEXT },
        { id: 'internalid', label: 'Item ID', type: ui.FieldType.TEXT },
        { id: 'displayname', label: 'Display Name', type: ui.FieldType.TEXT },
        { id: 'custitem_mi_cr_itm_cat', label: 'Core Item Category', type: ui.FieldType.TEXT },
        { id: 'custitem_mi_product_category', label: 'Product Category', type: ui.FieldType.TEXT },
        { id: 'cseg_mi_brand', label: 'Brand', type: ui.FieldType.TEXT }
      ].forEach(function (f) {
        sublist.addField({ id: f.id, label: f.label, type: f.type });
        sublistFieldMeta[f.id] = f.label;
      });

      var priceLevelList = Object.keys(priceLevels);
      priceLevelList.forEach(function (plName) {
        var colId = plName.toLowerCase().replace(/[^a-z0-9]/gi, '_');
        sublist.addField({ id: colId, label: plName, type: ui.FieldType.CURRENCY });
        priceLevels[plName] = colId;
        sublistFieldMeta[colId] = plName;
      });

      // Fill rows
      var line = 0;
      for (var id in itemSearchResults) {
        var data = itemSearchResults[id];
        // static
        ['itemid','internalid','displayname','custitem_mi_cr_itm_cat','custitem_mi_product_category','cseg_mi_brand']
          .forEach(function(k){
            if (data[k]) {
              sublist.setSublistValue({ id: k, line: line, value: String(data[k]) });
            }
          });
        // dynamic prices
        for (var pl in data.prices) {
          var priceVal = data.prices[pl];
          if (priceVal) {
            sublist.setSublistValue({ id: priceLevels[pl], line: line, value: String(priceVal) });
          }
        }
        line++;
      }

      // Store column ids + meta for CSV
      var columnIdList = ['itemid', 'internalid', 'displayname', 'salesdescription', 'custitem_mi_cr_itm_cat', 'custitem_mi_product_category', 'cseg_mi_brand']
        .concat(priceLevelList.map(function(plName){ return plName.toLowerCase().replace(/[^a-z0-9]/gi, '_'); }));
      var colField = form.addField({ id: 'custpage_all_column_ids', label: 'All Column IDs', type: ui.FieldType.LONGTEXT });
      colField.defaultValue = JSON.stringify(columnIdList);
      colField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

      var metaField = form.addField({ id: 'custpage_sublist_fieldmeta', label: 'Sublist Field Metadata', type: ui.FieldType.LONGTEXT });
      metaField.defaultValue = JSON.stringify(sublistFieldMeta);
      metaField.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });
      var resetParam = {
        e: empid,
        t: ts,
        s: sig
      }

      form.addButton({ id: 'reset', label: 'Reset', functionName: 'onReset('+ JSON.stringify(resetParam) + ')' });

      // Client script
      form.clientScriptFileId = ClientID;

      context.response.writePage(form);
    }
    else if (context.request.method === 'POST') {
      try {
        var sublistId = 'custpage_itemlist';
        var lineCount = context.request.getLineCount({ group: sublistId });

        // Read stored ID-label mapping
        var fieldMeta = JSON.parse(context.request.parameters.custpage_sublist_fieldmeta || '{}');
        var fieldIds = Object.keys(fieldMeta);

        // CSV header
        var csvHeader = fieldIds.map(function (fieldId) {
          return '"' + (fieldMeta[fieldId] || fieldId) + '"';
        }).join(',') + '\n';

        // Rows
        var csvRows = '';
        for (var i = 0; i < lineCount; i++) {
          var row = fieldIds.map(function (fieldId) {
            var value = context.request.getSublistValue({ group: sublistId, name: fieldId, line: i });
            if (value == null) return '""';
            value = String(value).replace(/"/g, '""');
            return '"' + value + '"';
          }).join(',');
          csvRows += row + '\n';
        }

        var csvContent = csvHeader + csvRows;

        var fileObj = file.create({
          name: 'Item_Inventory_Report.csv',
          fileType: file.Type.CSV,
          contents: csvContent,
          folder: 363575
        });

        var fileId = fileObj.save();
        var fileObj1 = file.load({ id: fileId });
        context.response.writeFile(fileObj1, false);

      } catch (e) {
        log.error({ title: 'Error Processing Form Submission', details: e });
        context.response.write('Error: ' + e.message);
      }
    }
  }

  return { onRequest: onRequest };
});
