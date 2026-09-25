/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/ui/serverWidget', 'N/search', 'N/runtime', 'N/record', 'N/crypto'], function (serverWidget, search, runtime, record, crypto) {
  function onRequest(context) {

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

      const form = serverWidget.createForm({ title: 'Planned PO Approval' });

      var fEmp = form.addField({ id: 'custpage_empid', label: 'empid', type: serverWidget.FieldType.TEXT });
      fEmp.defaultValue = empid; fEmp.updateDisplayType({ displayType: serverWidget.FieldDisplayType.HIDDEN });
      var fTs = form.addField({ id: 'custpage_ts', label: 'ts', type: serverWidget.FieldType.TEXT });
      fTs.defaultValue = ts; fTs.updateDisplayType({ displayType: serverWidget.FieldDisplayType.HIDDEN });
      var fSig = form.addField({ id: 'custpage_sig', label: 'sig', type: serverWidget.FieldType.TEXT });
      fSig.defaultValue = sig; fSig.updateDisplayType({ displayType: serverWidget.FieldDisplayType.HIDDEN });

      const htmlField = form.addField({
        id: 'custpage_html_field',
        type: serverWidget.FieldType.INLINEHTML,
        label: 'Planned PO Interface'
      });

      // Only approval data; filter where approval status is NULL/empty
      const approvalResults = getApprovalResultsOnly();
      const vendorsList = getAllVendors();

      // ⬇️ Pass the signed session values into the HTML so the POST includes them
      htmlField.defaultValue = generateApprovalOnlyHtml(approvalResults, vendorsList, empid, ts, sig);

      context.response.writePage(form);
    }

    if (context.request.method === 'POST') {
      const params = context.request.parameters;
      const selectedIds = (params.custpage_selected_ids || '').split(',').filter(Boolean);
      const action = params.custpage_action_type;

      // token fields posted from the HTML form
      var postedEmp = params.custpage_empid || '';
      var postedTs  = params.custpage_ts || '';
      var postedSig = params.custpage_sig || '';
      var authorized = (postedEmp && postedTs && postedSig && verify(postedEmp, postedTs, postedSig));

      // --- Persist inline edits for SELECTED rows (only if token is valid) ---
// --- Handle actions ---
if (action === 'save' && authorized) {
  // Changed IDs come from client (tracked via data-dirty)
  const changedIds = (params.custpage_changed_ids || '').split(',').filter(Boolean);

  if (!changedIds.length) {
    // no-op; just bounce back
  } else {
    changedIds.forEach(function (id) {
      try {
        var qtyStr  = params['qty_'   + id];
        var costStr = params['cost_'  + id];
        var vendVal = params['vendor_'+ id];
        var memoStr = params['memo_' + id];

        var values = {};

        if (qtyStr !== undefined && qtyStr !== '') {
          var q = Math.floor(Number(qtyStr));
          if (Number.isFinite(q) && q >= 0) values.custrecord_mi_order_qty = String(q);
        }
        if (costStr !== undefined && costStr !== '') {
          var c = Number(costStr);
          if (Number.isFinite(c) && c >= 0) values.custrecord_vendor_rate = c.toFixed(2);
        }
        if (vendVal !== undefined && vendVal !== '') {
          values.custrecord_mi_preffered_vendor = vendVal;
        }
        if (memoStr !== undefined) {
          values.custrecord_mi_purchase_memo = String(memoStr || '');
        }

        if (Object.keys(values).length) {
          record.submitFields({
            type: 'customrecord_mi_planned_po',
            id: id,
            values: values
          });
        }
      } catch (e) {
        log.error('Save failed for ' + id, e);
      }
    });
  }
}



      if (selectedIds.length && action) {
        if (action === 'approve' || action === 'reject') {
          const newStatus = action === 'approve' ? '1' : '3';
          selectedIds.forEach(id => {
            try {
              record.submitFields({
                type: 'customrecord_mi_planned_po',
                id: id,
                values: { custrecord_mi_approval_status: newStatus }
              });
            } catch (e) {
              log.error('Error updating status for ID ' + id, e);
            }
          });
        }
      }

      // back to the same Suitelet with the signed params
      context.response.write(
        '<script>window.location.href = "https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2520&deploy=1&compid=4975346&ns-at=AAEJ7tMQVziJF1qjHHrnbuPycJ3uz76kG7ifpvym_jIq-Whr50U&empid=' +
        encodeURIComponent(postedEmp) + '&ts=' + encodeURIComponent(postedTs) + '&sig=' + encodeURIComponent(postedSig) + '";</script>'
      );
    }
  }

  // --- SEARCH: approval status is NULL/empty only ---
  function getApprovalResultsOnly() {
    var results = [];
    const searchObj = search.create({
      type: 'customrecord_mi_planned_po',
      filters: [
        ['custrecord_mi_approval_status', 'anyof', '@NONE@']
      ],
      columns: [
        'internalid',
        'custrecord_mi_item',
        'custrecord_item_name',
        'custrecord_vendor_rate',
        'custrecord_vendor_currency',
        'custrecord_mi_preffered_vendor',
        'custrecord_mi_order_qty',
        'custrecord_mi_preffered_vendor_min_qty',
        'custrecord_mi_qty_available',
        'custrecord_mi_qty_in_transit',
        'custrecord_mi_qty_on_order_not_recv',
        'custrecord_mi_qty_back_ordered',
        'custrecord_mi_min_month_qty',
        'custrecord_mi_purchase_memo',
        'custrecord_month_of_stocks'
      ]
    });

    searchObj.run().each(result => {
      results.push({
        id: result.id,
        item: result.getText('custrecord_mi_item'),
        itemId: result.getValue('custrecord_mi_item'),
        itemName: result.getValue('custrecord_item_name'),
        cost: result.getValue('custrecord_vendor_rate'),
        currency: result.getText('custrecord_vendor_currency'),
        vendor: result.getText('custrecord_mi_preffered_vendor'),
        vendorId: result.getValue('custrecord_mi_preffered_vendor'),
        orderQty: result.getValue('custrecord_mi_order_qty'),
        minQty: result.getValue('custrecord_mi_preffered_vendor_min_qty'),
        available: result.getValue('custrecord_mi_qty_available'),
        inTransit: result.getValue('custrecord_mi_qty_in_transit'),
        onOrder: result.getValue('custrecord_mi_qty_on_order_not_recv'),
        backOrdered: result.getValue('custrecord_mi_qty_back_ordered'),
        minMonthQty: result.getValue('custrecord_mi_min_month_qty'),
        memo: result.getValue('custrecord_mi_purchase_memo'),
        mstock: result.getValue('custrecord_month_of_stocks')
      });
      return true;
    });

    return results;
  }
  function getAllVendors() {
  var return_vendors = [];
  var vendorSearch = search.create({
    type: search.Type.VENDOR,
    filters: [
      ['isinactive', 'is', 'F']
    ],
    columns: [
      'internalid',
      'altname'
    ]
  });
    return_vendors.push({
      id: '',
      name: ''
    });
  vendorSearch.run().each(function(res){
    return_vendors.push({
      id: res.getValue('internalid'),
      name: res.getValue('altname')
    });
    return true;
  });
  return return_vendors;
}

  // --- UI: approval subtab only (no Create PO) ---
  function generateApprovalOnlyHtml(data, vendors, empid, ts, sig) {
    return `
<style>
  select[multiple] {
    width: 220px;
    height: 80px;
    padding: 6px;
    border-radius: 4px;
    border: 1px solid #ccc;
    margin-right: 10px;
  }
  label {
    font-weight: 600;
    display: block;
    margin: 10px 0 4px;
  }

  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }
  table th,
  table td {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: center;
  }
  table th {
    background-color: #f8f8f8;
    font-weight: bold;
  }

  /* row striping, skipped if dirty */
  table tr:nth-child(even):not(.dirty) {
    background-color: #f9f9f9;
  }

  /* highlight dirty rows (always wins) */
  table tr.dirty {
    background-color: #fff7e6;
  }

  table input[type="number"],
  table select {
    width: 120px;
    padding: 4px 6px;
  }
  table select.vendSel {
    width: 250px;
    max-width: 100%;
  }

  button {
    background-color: #007bff;
    color: #fff;
    border: none;
    padding: 8px 16px;
    margin: 10px 5px 20px 0;
    border-radius: 4px;
    font-size: 13px;
    cursor: pointer;
  }
  button:hover {
    background-color: #0056b3;
  }

  .filter-row {
    display: flex;
    flex-wrap: wrap;
    gap: 20px;
    margin: 10px 0 20px;
  }

  h2 {
    margin: 0 0 10px 0;
  }
</style>


      <h2>Approval Action</h2>
      ${generateApprovalContent(data, vendors, empid, ts, sig)}
    `;
  }

  function generateApprovalContent(data, vendors, empid, ts, sig) {
    const vendorOptions = vendors
    .map(v => `<option value="${v.id}">${v.name}</option>`)
    .join('');
    var vendors = [...new Set(data.map(row => `<option value="${row.vendorId}">${row.vendor}</option>`))].join('');
    const items = [...new Set(data.map(row => `<option value="${row.itemId}">${row.item}</option>`))].join('');

    return `
      <form method="POST">
        <input type="hidden" name="custpage_selected_ids" id="custpage_selected_ids_approval" />
        <input type="hidden" name="custpage_action_type" id="custpage_action_type_approval" />

        <!-- Post signed session back with the form -->
        <input type="hidden" name="custpage_empid" value="${String(empid || '').replace(/"/g,'&quot;')}" />
        <input type="hidden" name="custpage_ts"    value="${String(ts || '').replace(/"/g,'&quot;')}" />
        <input type="hidden" name="custpage_sig"   value="${String(sig || '').replace(/"/g,'&quot;')}" />
        <input type="hidden" name="custpage_changed_ids" id="custpage_changed_ids" />
        <input type="hidden" name="custpage_has_dirty" id="custpage_has_dirty" value="0" />


        <div class="filter-row">
          <div>
            <label>Filter by Vendor:</label>
            <select id="vendorFilter_approval" multiple>${vendors}</select>
          </div>
          <div>
            <label>Filter by Item:</label>
            <select id="itemFilter_approval" multiple>${items}</select>
          </div>
        </div>

        <div>
          <button type="submit" onclick="return submitUpdates()">Submit Updates</button>
          <button type="submit" onclick="return setAction('approval','approve')">Approve</button>
          <button type="submit" onclick="return setAction('approval','reject')">Reject</button>
        </div>

        <table>
          <thead>
            <tr>
              <th><input type="checkbox" id="checkAll_approval" onclick="toggleAll(this,'approval')" /></th>
              <th>Item</th><th>Item Id</th><th>Display Name</th><th>Purchase Memo</th><th>Vendor</th><th>Ordered Qty</th><th>Months of Stock</th><th>Vendor Price</th><th>Currency</th><th>Vendor Min Qty</th>
              <th>Not Shipped</th><th>Available</th><th>In Transit</th><th>On Order</th><th>Min Month Qty</th>
            </tr>
          </thead>
<tbody id="tableBody_approval">
  ${data.map(row => `
    <tr data-id="${row.id}" data-item="${row.itemId}" data-vendor="${row.vendorId}">
  <td><input type="checkbox" class="selectLine_approval" name="selectLine_approval" /></td>
  <td>${row.item}</td>
  <td>${row.itemId}</td>
  <td>${row.itemName}</td>
<td>
  <input
    type="text"
    name="memo_${row.id}"
    class="memoInput"
    value="${(row.memo || '').replace(/"/g, '&quot;')}"
    data-original="${(row.memo || '').replace(/"/g, '&quot;')}"
    oninput="markDirty(this)"
    onchange="markDirty(this)"
  />
</td>
  <td>
    <select name="vendor_${row.id}" class="vendSel" 
            data-original="${row.vendorId}">
      ${vendorOptions.replace(
        `value="${row.vendorId}"`,
        `value="${row.vendorId}" selected`
      )}
    </select>
  </td>
  <td>
    <input
      type="number"
      name="qty_${row.id}"
      class="qtyInput"
      value="${row.orderQty || ''}"
      data-original="${row.orderQty || ''}"
      min="0"
      step="1"
      inputmode="numeric"
      pattern="\\d*"
      oninput="this.value = this.value.replace(/[^0-9]/g,''); markDirty(this)"
      onchange="markDirty(this)"
    />
  </td>
  <td>${row.mstock}</td>
  <td>
    <input
      type="number"
      name="cost_${row.id}"
      class="costInput"
      value="${row.cost || ''}"
      data-original="${row.cost || ''}"
      min="0"
      step="0.01"
      inputmode="decimal"
      oninput="this.value = this.value.replace(/[^0-9.]/g,'').replace(/(\\..*)\\./g,'$1'); markDirty(this)"
      onchange="markDirty(this)"
    />
  </td>
  <td>${row.currency}</td>
  <td>${row.minQty}</td>
  <td>${row.backOrdered}</td>
  <td>${row.available}</td>
  <td>${row.inTransit}</td>
  <td>${row.onOrder}</td>
  <td>${row.minMonthQty}</td>
</tr>
  `).join('')}
</tbody>
        </table>

        <script>
        function rowEl(el){ return el.closest('tr'); }


function markDirty(inputOrSelect){
  const tr = inputOrSelect.closest('tr');
  if (!tr) return;

  const vend = tr.querySelector('select.vendSel');
  const qty  = tr.querySelector('input.qtyInput');
  const cost = tr.querySelector('input.costInput');
  const memo = tr.querySelector('input.memoInput');

  const vendDirty = vend && vend.value !== (vend.getAttribute('data-original') || '');
  const qtyDirty  = qty  && (qty.value  !== (qty.getAttribute('data-original') || ''));
  const costDirty = cost && (cost.value !== (cost.getAttribute('data-original') || ''));
  const memoDirty = memo && (memo.value !== (memo.getAttribute('data-original') || ''));

  const isDirtyRow = !!(vendDirty || qtyDirty || costDirty || memoDirty);
  tr.classList.toggle('dirty', isDirtyRow);

  setDirtyFlag(hasDirty());
}


function collectChangedIds(){
  return Array.from(document.querySelectorAll('#tableBody_approval tr.dirty'))
    .map(tr => tr.dataset.id);
}

function collectSelectedIds(type){
  return Array.from(document.querySelectorAll('#tableBody_'+type+' tr:not([style*="display: none"]) .selectLine_'+type+':checked'))
    .map(cb => cb.closest('tr').dataset.id);
}

function setDirtyFlag(isDirty){
  var f = document.getElementById('custpage_has_dirty');
  if (f) f.value = isDirty ? '1' : '0';
}

function hasDirty(){
  return document.querySelector('#tableBody_approval tr.dirty') !== null;
}

function submitUpdates(){
  const changed = Array.from(document.querySelectorAll('#tableBody_approval tr.dirty'))
    .map(tr => tr.dataset.id);

  if (!changed.length){
    alert('Please update at least one line before submitting.');
    return false;
  }
  document.getElementById('custpage_changed_ids').value = changed.join(',');
  document.getElementById('custpage_action_type_approval').value = 'save';
  return true; // allow POST
}

// --- Approve/Reject with guards ---
window.setAction = function(type, action){
  // 1. Check for unsaved changes first
  if (hasDirty()){
    alert('You have unsaved changes. Please click "Submit Updates" first, then ' + action + '.');
    return false;
  }

  // 2. Collect selected IDs
  const selectedIds = Array.from(
    document.querySelectorAll('#tableBody_'+type+' tr:not([style*="display: none"]) .selectLine_'+type+':checked')
  ).map(cb => cb.closest('tr').dataset.id);

  // 3. Require at least one line selected
  if(!selectedIds.length){
    alert('Please select at least one line to perform this action.');
    return false;
  }

  // 4. Pass values into hidden fields for POST
  document.getElementById('custpage_selected_ids_'+type).value = selectedIds.join(',');
  document.getElementById('custpage_action_type_'+type).value = action;
  return true;
};



// initialize change listener for vendor selects
document.addEventListener('DOMContentLoaded', function(){
  document.querySelectorAll('select.vendSel').forEach(function(sel){
    sel.addEventListener('change', function(){ markDirty(sel); });
  });
});

          window.toggleAll = function(src, type){
            const visibleRows = document.querySelectorAll('#tableBody_'+type+' tr:not([style*="display: none"])');
            visibleRows.forEach(row => {
              const cb = row.querySelector('.selectLine_'+type);
              if(cb) cb.checked = src.checked;
            });
          };
          window.filterTable = function(type){
            const vendorValues = Array.from(document.getElementById('vendorFilter_'+type).selectedOptions).map(o=>o.value);
            const itemValues = Array.from(document.getElementById('itemFilter_'+type).selectedOptions).map(o=>o.value);
            document.querySelectorAll('#tableBody_'+type+' tr').forEach(row => {
              const show = (!vendorValues.length || vendorValues.includes(row.dataset.vendor)) &&
                           (!itemValues.length || itemValues.includes(row.dataset.item));
              row.style.display = show ? '' : 'none';
            });
          };
          document.addEventListener('DOMContentLoaded', function(){
            document.getElementById('vendorFilter_approval').addEventListener('change', function(){ filterTable('approval'); });
            document.getElementById('itemFilter_approval').addEventListener('change', function(){ filterTable('approval'); });
            const checkAllBox = document.getElementById('checkAll_approval');
            if(checkAllBox){ checkAllBox.addEventListener('change', function(){ toggleAll(this,'approval'); }); }
          });
        </script>
      </form>
    `;
  }

  return { onRequest };
});
