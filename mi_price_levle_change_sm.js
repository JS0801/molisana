/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/ui/serverWidget', 'N/search', 'N/record', 'N/runtime', 'N/log', 'N/format', 'N/url', 'N/redirect', 'N/crypto'],
function (ui, search, record, runtime, log, format, url, redirect, crypto) {

  function onRequest(context) {

    // --- Portal URL used for login bounce ---
    var portalUrl = 'https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2110&deploy=1&compid=4975346&ns-at=AAEJ7tMQamzukv1WMqTK6i2c27bRetbrd2MDLjhDgPPFOawMxCo';

    // --- Signed session helpers (same as other tools) ---
    const SECRET = runtime.getCurrentScript().getParameter({ name: 'custscript_portal_secret' }) || 'change-me';
    const TOKEN_TTL_MS = 30 * 60 * 1000; // 30 minutes

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
      var form = ui.createForm({ title: 'Price Change Tool' });


            var q = context.request.parameters || {};
      var empid = q.empid || '';
      var ts = q.ts || '';
      var sig = q.sig || '';

      // require a valid signed session
      if (!(empid && ts && sig && verify(empid, ts, sig))) {
        context.response.write(
          '<html><head>' +
          '<script>setTimeout(function(){ window.location.href = ' + JSON.stringify(portalUrl) + '; }, 1200);</script>' +
          '<style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>' +
          '</head><body><div class="message">Login Required</div></body></html>'
        );
        return;
      }


      var fEmp = form.addField({ id: 'custpage_empid', label: 'empid', type: ui.FieldType.TEXT });
      fEmp.defaultValue = empid; fEmp.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });
      var fTs = form.addField({ id: 'custpage_ts', label: 'ts', type: ui.FieldType.TEXT });
      fTs.defaultValue = ts; fTs.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });
      var fSig = form.addField({ id: 'custpage_sig', label: 'sig', type: ui.FieldType.TEXT });
      fSig.defaultValue = sig; fSig.updateDisplayType({ displayType: ui.FieldDisplayType.HIDDEN });

      // Custom Inline Style
      form.addField({
        id: 'custpage_html_style',
        label: 'Style',
        type: ui.FieldType.INLINEHTML
      }).defaultValue = `
        <style>
          .ns-form .uir-field-group { margin-bottom: 24px; }
          .ns-form .uir-field-wrapper { padding: 4px 0; }
          .ns-form .uir-field-label { font-weight: bold; }
        </style>
      `;

      // Field Groups
      form.addFieldGroup({ id: 'custpage_filter_group', label: 'Filter Criteria' });
      form.addFieldGroup({ id: 'custpage_customer_group', label: 'Customer Selection & Date Range' });
      form.addFieldGroup({ id: 'custpage_behavior_group', label: 'Settings' });

      // Filter Fields
      var coreCategoryField = form.addField({
        id: 'custpage_core_category',
        label: 'Core Item Category',
        type: ui.FieldType.MULTISELECT,
        source: '876',
        container: 'custpage_filter_group'
      });

      var brandField = form.addField({
        id: 'custpage_brand',
        label: 'Brand',
        type: ui.FieldType.MULTISELECT,
        source: '570',
        container: 'custpage_filter_group'
      });
      brandField.updateBreakType({ breakType: ui.FieldBreakType.STARTCOL });

      var prodCategoryField = form.addField({
        id: 'custpage_prod_category',
        label: 'Product Category',
        type: ui.FieldType.MULTISELECT,
        source: '675',
        container: 'custpage_filter_group'
      });
      prodCategoryField.updateBreakType({ breakType: ui.FieldBreakType.STARTCOL });

      // Customer Fields
      var customerField = form.addField({
        id: 'custpage_customers',
        label: 'Customers',
        type: ui.FieldType.MULTISELECT,
        source: 'customer',
        container: 'custpage_customer_group'
      });
      customerField.isMandatory = true;

      var startDateField = form.addField({
        id: 'custpage_startdate',
        label: 'Start Date',
        type: ui.FieldType.DATE,
        container: 'custpage_customer_group'
      });
      startDateField.isMandatory = true;
      startDateField.updateBreakType({ breakType: ui.FieldBreakType.STARTCOL });

      var endDateField = form.addField({
        id: 'custpage_enddate',
        label: 'End Date',
        type: ui.FieldType.DATE,
        container: 'custpage_customer_group'
      });
      endDateField.isMandatory = true;

      var applyToCustomerField = form.addField({
        id: 'custpage_apply_to_customers',
        label: 'Apply to Customers',
        type: ui.FieldType.CHECKBOX,
        container: 'custpage_customer_group'
      });
      applyToCustomerField.updateBreakType({ breakType: ui.FieldBreakType.STARTCOL });

      // Default Filters
      var reqParams = context.request.parameters;
      var coreCat = (reqParams.custpage_core_category || '').split(',').filter(Boolean);
      var brand = (reqParams.custpage_brand || '').split(',').filter(Boolean);
      var prodCat = (reqParams.custpage_prod_category || '').split(',').filter(Boolean);

      coreCategoryField.defaultValue = coreCat;
      brandField.defaultValue = brand;
      prodCategoryField.defaultValue = prodCat;

      // Sublist
      var sublist = form.addSublist({
        id: 'custpage_items',
        label: 'Item List',
        type: ui.SublistType.LIST
      });

      sublist.addField({ id: 'itemid', label: 'Item Name', type: ui.FieldType.TEXT });
      sublist.addField({ id: 'internalid', label: 'Item ID', type: ui.FieldType.TEXT });
      sublist.addField({ id: 'displayname', label: 'Display Name', type: ui.FieldType.TEXT });
      sublist.addField({ id: 'corecat', label: 'Core Category', type: ui.FieldType.TEXT });
      sublist.addField({ id: 'prodcat', label: 'Product Category', type: ui.FieldType.TEXT });
      sublist.addField({ id: 'brand', label: 'Brand', type: ui.FieldType.TEXT });
      sublist.addField({ id: 'baseprice', label: 'Green Price (Base)', type: ui.FieldType.CURRENCY });
      sublist.addField({ id: 'newprice', label: 'New Price', type: ui.FieldType.CURRENCY })
        .updateDisplayType({ displayType: ui.FieldDisplayType.ENTRY });

      // Item Search Filters
      var itemFilters = [['isinactive', 'is', 'F']];
      if (coreCat.length) itemFilters.push('AND', ['custitem_mi_cr_itm_cat', 'anyof'].concat(coreCat));
      if (brand.length) itemFilters.push('AND', ['cseg_mi_brand', 'anyof'].concat(brand));
      if (prodCat.length) itemFilters.push('AND', ['custitem_mi_product_category', 'anyof'].concat(prodCat));

      var resultSet = search.create({
        type: 'item',
        filters: itemFilters,
        columns: ['itemid', 'internalid', 'displayname', 'custitem_mi_cr_itm_cat', 'custitem_mi_product_category', 'cseg_mi_brand', 'baseprice']
      }).run().getRange({ start: 0, end: 1000 });

      for (var i = 0; i < resultSet.length; i++) {
        var res = resultSet[i];
        sublist.setSublistValue({ id: 'itemid', line: i, value: res.getValue('itemid') || ' ' });
        sublist.setSublistValue({ id: 'internalid', line: i, value: res.getValue('internalid') });
        sublist.setSublistValue({ id: 'displayname', line: i, value: res.getValue('displayname') || ' ' });
        sublist.setSublistValue({ id: 'corecat', line: i, value: res.getText('custitem_mi_cr_itm_cat') || ' ' });
        sublist.setSublistValue({ id: 'prodcat', line: i, value: res.getText('custitem_mi_product_category') || ' ' });
        sublist.setSublistValue({ id: 'brand', line: i, value: res.getText('cseg_mi_brand') || ' ' });
        sublist.setSublistValue({ id: 'baseprice', line: i, value: res.getValue('baseprice') || '0' });
      }

      // Buttons
      form.addSubmitButton('Save Price Changes');
var fn = 'reset('
  + [empid, ts, sig].map(function (v) { return JSON.stringify(String(v || '')); }).join(', ')
  + ')';

form.addButton({
  id: 'custpage_refresh_btn',
  label: 'Reset',
  functionName: fn
});
      var clientScriptPath = runtime.getCurrentScript().getParameter({
        name: 'custscript_mi__price_tool_client'
      });
      if (clientScriptPath) form.clientScriptFileId = clientScriptPath;

      context.response.writePage(form);
    }

    // === POST Processing ===
    else {
      var req = context.request;
      var customers = req.parameters.custpage_customers;
      var startDateVal = req.parameters.custpage_startdate;
      var endDateVal = req.parameters.custpage_enddate;
      var applyToCustomer = req.parameters.custpage_apply_to_customers === 'T';
      var postedEmp = req.parameters.custpage_empid || '';
      var postedTs  = req.parameters.custpage_ts || '';
      var postedSig = req.parameters.custpage_sig || '';
      var authorized = (postedEmp && postedTs && postedSig && verify(postedEmp, postedTs, postedSig));


      
      if (!authorized) {
        context.response.write(
          '<html><head>' +
          '<script>setTimeout(function(){ window.location.href = ' + JSON.stringify(portalUrl) + '; }, 1200);</script>' +
          '<style>body{display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial;background:#0b0b0b;color:#fff}.message{font-size:20px;font-weight:700}</style>' +
          '</head><body><div class="message">Session expired. Please log in again.</div></body></html>'
        );
        return;
      }


      if (!customers || !startDateVal || !endDateVal) {
        throw 'Customers, Start Date, and End Date are required.';
      }

      var parsedStartDate = format.parse({ value: startDateVal, type: format.Type.DATE });
      var parsedEndDate = format.parse({ value: endDateVal, type: format.Type.DATE });
      var customerArray = customers.split(',');
      var lineCount = req.getLineCount({ group: 'custpage_items' });

      for (var i = 0; i < lineCount; i++) {
        var itemId = req.getSublistValue({ group: 'custpage_items', name: 'internalid', line: i });
        var oldPrice = req.getSublistValue({ group: 'custpage_items', name: 'baseprice', line: i });
        var newPrice = req.getSublistValue({ group: 'custpage_items', name: 'newprice', line: i });

        var oldPriceNum = parseFloat(oldPrice || 0);
        var newPriceNum = parseFloat(newPrice || 0);

        if (itemId && newPrice && newPriceNum !== oldPriceNum) {
          record.create({
            type: 'customrecord_ds_price_level_change_reque',
            isDynamic: true
          })
            .setValue('custrecord_ds_price_level_item', itemId)
            .setValue('custrecord_ds_old_price', oldPriceNum)
            .setValue('custrecord_ds_new_price', newPriceNum)
            .setValue('custrecord_ds_start_date', parsedStartDate)
            .setValue('custrecord_ds_end_date', parsedEndDate)
            .setValue('custrecord_mi_created_by_sm', true)
            .setValue({
              fieldId: 'custrecord_mi_customers',
              value: customerArray
            })
            .save();
            var cusids =  customerArray[0].split("\u0005")
          if (applyToCustomer) {
            log.debug('customerArray',cusids)
            for (var c = 0; c < cusids.length; c++) {
              var custId = cusids[c];
              var custRec = record.load({
                type: 'customer',
                id: custId,
                isDynamic: true
              });

              var itemPricingLineCount = custRec.getLineCount({ sublistId: 'itempricing' });
              
             var shouldAddNewLine = true;

for (var j = itemPricingLineCount - 1; j >= 0; j--) {
  custRec.selectLine({ sublistId: 'itempricing', line: j });

  var existingItem = custRec.getCurrentSublistValue({
    sublistId: 'itempricing',
    fieldId: 'item'
  });

  var existingPrice = custRec.getCurrentSublistValue({
    sublistId: 'itempricing',
    fieldId: 'price'
  });

  var existingLevel = custRec.getCurrentSublistValue({
    sublistId: 'itempricing',
    fieldId: 'level'
  });

  var existingCurrency = custRec.getCurrentSublistValue({
    sublistId: 'itempricing',
    fieldId: 'currency'
  });

  if (parseInt(existingItem) === parseInt(itemId) &&
      parseFloat(existingPrice) === newPriceNum &&
      parseInt(existingLevel) === -1 &&
      parseInt(existingCurrency) === 1) {
    // Exact match exists, skip update
    shouldAddNewLine = false;
    break;
  }

  if (parseInt(existingItem) === parseInt(itemId) &&
      parseInt(existingLevel) === -1 &&
      parseInt(existingCurrency) === 1) {
    // Price mismatch — remove outdated line
    custRec.removeLine({ sublistId: 'itempricing', line: j, ignoreRecalc: true });
  }
}

if (shouldAddNewLine) {
  custRec.selectNewLine({ sublistId: 'itempricing' });
  custRec.setCurrentSublistValue({ sublistId: 'itempricing', fieldId: 'item', value: itemId });
  custRec.setCurrentSublistValue({ sublistId: 'itempricing', fieldId: 'level', value: -1 });
  custRec.setCurrentSublistValue({ sublistId: 'itempricing', fieldId: 'currency', value: 1 });
  custRec.setCurrentSublistValue({ sublistId: 'itempricing', fieldId: 'price', value: newPriceNum });
  custRec.commitLine({ sublistId: 'itempricing' });
}


              custRec.save();
            }
          }
        }
      }

context.response.write(
        '<html><body>' +
        '<script>' +
        'window.location.href=' +
        JSON.stringify('https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2311&deploy=1&compid=4975346&ns-at=AAEJ7tMQL9UOA94WkZJa9hx-tE5vj4RGzOoM-JVdXqNbMHypnes' +
                       '&empid=' + encodeURIComponent(postedEmp) +
                       '&ts='    + encodeURIComponent(postedTs) +
                       '&sig='   + encodeURIComponent(postedSig)) +
        ';' +
        '</script>' +
        '</body></html>'
      );
    }
  }

  return { onRequest: onRequest };
});
