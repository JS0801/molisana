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

            const approvalResults = getSearchResults('approvalaction');
            const createPoResults = getSearchResults('createpo');
            const vendorsList = getAllVendors();

            htmlField.defaultValue = generateTabsHtml(approvalResults, createPoResults, vendorsList, empid, ts, sig);

            context.response.writePage(form);
        }

        if (context.request.method === 'POST') {
            const params = context.request.parameters;
            const selectedIds = (params.custpage_selected_ids || '').split(',').filter(Boolean);
            const action = params.custpage_action_type;



            // Verify token posted from hidden fields
            var postedEmp = params.custpage_empid || '';
            var postedTs = params.custpage_ts || '';
            var postedSig = params.custpage_sig || '';
            var authorized = (postedEmp && postedTs && postedSig && verify(postedEmp, postedTs, postedSig));

            if (action === 'save' && authorized) {
                const changedIds = (params.custpage_changed_ids || '').split(',').filter(Boolean);
                if (changedIds.length) {
                    changedIds.forEach(function (id) {
                        try {
                            var qtyStr = params['qty_' + id];
                            var costStr = params['cost_' + id];
                            var vendVal = params['vendor_' + id];
                            var memoStr = params['memo_' + id];
                            var values = {};

                            if (memoStr !== undefined) {
                                values.custrecord_mi_purchase_memo = memoStr;
                            }
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
                            if (Object.keys(values).length) {
                                record.submitFields({
                                    type: 'customrecord_mi_planned_po',
                                    id: id,
                                    values: values,
                                    options: { enableSourcing: true, ignoreMandatoryFields: true }
                                });
                            }
                        } catch (e) { log.error('Save failed for ' + id, e); }
                    });
                }
            }

            if ((action === 'approve' || action === 'reject' || action === 'createpo') && (!selectedIds || selectedIds.length === 0)) {
                context.response.write(
                    '<script>alert("Please select at least one line."); window.history.back();</script>'
                );
                return;
            }

            if (selectedIds.length && action) {
                if (action === 'approve' || action === 'reject') {
                    const newStatus = action === 'approve' ? '2' : '3';
                    selectedIds.forEach(id => {
                        try {
                            record.submitFields({
                                type: 'customrecord_mi_planned_po',
                                id: id,
                                values: { custrecord_mi_approval_status: newStatus }
                            });
                        } catch (e) {
                            log.error(`Error updating status for ID ${id}`, e);
                        }
                    });
                }

                if (action === 'createpo') {
                    const vendorMap = {};
                    const recordVendorMap = {};

                    selectedIds.forEach(id => {
                        try {
                            const rec = record.load({ type: 'customrecord_mi_planned_po', id });
                            const vendorId = rec.getValue('custrecord_mi_preffered_vendor');
                            const itemId = rec.getValue('custrecord_mi_item');
                            const qty = rec.getValue('custrecord_mi_order_qty');
                            const memo = rec.getValue('custrecord_mi_purchase_memo');

                            if (!vendorMap[vendorId]) vendorMap[vendorId] = [];
                            vendorMap[vendorId].push({ itemId, qty, memo });

                            if (!recordVendorMap[vendorId]) recordVendorMap[vendorId] = [];
                            recordVendorMap[vendorId].push(id); // Save the planned PO record ID for later linking

                        } catch (e) {
                            log.error(`Error loading planned PO record ${id}`, e);
                        }
                    });

                    Object.keys(vendorMap).forEach(vendorId => {
                        try {
                            const po = record.create({ type: 'purchaseorder', isDynamic: true });
                            po.setValue({ fieldId: 'entity', value: vendorId });
                            po.setValue({ fieldId: 'custbody_planned_po', value: true });

                            vendorMap[vendorId].forEach(line => {
                                po.selectNewLine({ sublistId: 'item' });
                                po.setCurrentSublistValue({ sublistId: 'item', fieldId: 'item', value: line.itemId });
                                po.setCurrentSublistValue({ sublistId: 'item', fieldId: 'quantity', value: line.qty });
                                po.setCurrentSublistValue({ sublistId: 'item', fieldId: 'custcol_purchase_note', value: line.memo });
                                po.commitLine({ sublistId: 'item' });
                            });

                            const poId = po.save({
    enableSourcing: true,
    ignoreMandatoryFields: true
});
                            log.audit('PO Created', `PO ID ${poId} for Vendor ${vendorId}`);

                            // Link PO back to all related planned PO records
                            (recordVendorMap[vendorId] || []).forEach(plannedId => {
                                try {
                                    record.submitFields({
                                        type: 'customrecord_mi_planned_po',
                                        id: plannedId,
                                        values: { custrecord_mi_linked_purchase_order: poId }
                                    });
                                } catch (e) {
                                    log.error(`Error linking PO ${poId} to planned record ${plannedId}`, e);
                                }
                            });

                        } catch (e) {
                            log.error(`Error creating PO for vendor ${vendorId}`, e);
                        }
                    });
                }

            }

            if (action === 'scrappo') {

                log.debug('scrape po selected ids', selectedIds);
                selectedIds.forEach(id => {


                    record.delete({
                      type: 'customrecord_mi_planned_po',
                      id: id
                    });
                  
                    // const rec = record.load({ type: 'customrecord_mi_planned_po', id: id });
                    // rec.setValue({ fieldId: 'isinactive', value: true });
                    // rec.save();
                });

                context.response.write(
                '<script>window.location.href =' + portalUrl + ';</script>'
                );

                // return;
            }

            context.response.write(
                '<script>window.location.href = "https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2313&deploy=1&compid=4975346&ns-at=AAEJ7tMQw0WmH41Gs02OQcTCXn0VTpmqW6ShNtGIxBftDD0rD7c&empid=' + encodeURIComponent(postedEmp) + '&ts=' + encodeURIComponent(postedTs) + '&sig=' + encodeURIComponent(postedSig) + '";</script>'
            );
        }
    }

    function getSearchResults(status) {
        log.debug('status', status)
        var results = [];
        if (status == 'approvalaction') {
            var filter = [
                [['custrecord_mi_approval_status', 'anyof', "1"], "AND", ['isinactive', 'is', "F"]]
            ]
        } else {
            var filter = [
                [
                    ['custrecord_mi_approval_status', 'anyof', "2"],
                    "AND",
                    ['custrecord_mi_linked_purchase_order', 'anyof', "@NONE@"],
                    "AND",
                    ['isinactive', 'is', "F"]
                ]
            ]
        }


        const searchObj = search.create({
            type: 'customrecord_mi_planned_po',
            filters: filter,
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

    function generateTabsHtml(approvalData, createPoData, vendors, empid, ts, sig) {
        return `
      <style>
        .tab-header { display:flex; border-bottom:2px solid #007bff; margin-bottom:15px; }
        .tab-header div { padding:12px 24px; cursor:pointer; font-weight:bold; border-radius:6px 6px 0 0; background:#f1f1f1; margin-right:5px; transition:background .3s; }
        .tab-header div:hover { background:#e0e0e0; }
        .tab-header .active { background:#007bff; color:#fff; border-bottom:none; }
        .tab-content { display:none; }
        .tab-content.active { display:block; }
  
        select[multiple]{ width:220px; height:80px; padding:6px; border-radius:4px; border:1px solid #ccc; margin-right:10px; }
        label{ font-weight:600; display:block; margin:10px 0 4px; }
  
        table{ width:100%; border-collapse:collapse; font-size:13px; }
        table th, table td{ border:1px solid #ddd; padding:8px; text-align:center; }
        table th{ background-color:#f8f8f8; font-weight:bold; }
        table tr:nth-child(even):not(.dirty){ background-color:#f9f9f9; }
        table tr.dirty{ background-color:#fff7e6; }
  
        table input[type="number"], table select { width:120px; padding:4px 6px; }
        table select.vendSel { width:250px; max-width:100%; }
  
        button{ background-color:#007bff; color:#fff; border:none; padding:8px 16px; margin:10px 5px 20px 0; border-radius:4px; font-size:13px; cursor:pointer; }
        button:hover{ background-color:#0056b3; }
  
        .filter-row{ display:flex; flex-wrap:wrap; gap:20px; margin:10px 0 20px; }
      </style>
      <div class="tab-header">
        <div id="tab-approval" class="active" onclick="showTab('approval')">Approval Action</div>
        <div id="tab-createpo" onclick="showTab('createpo')">Create Purchase Order</div>
      </div>
      <div id="content-approval" class="tab-content active">
        ${generateTabContent(approvalData, 'approval', vendors, empid, ts, sig)}
      </div>
      <div id="content-createpo" class="tab-content">
        ${generateTabContent(createPoData, 'createpo', vendors, empid, ts, sig)}
      </div>
      <script>
        function showTab(tab){
          document.querySelectorAll('.tab-header div').forEach(el=>el.classList.remove('active'));
          document.querySelectorAll('.tab-content').forEach(el=>el.classList.remove('active'));
          document.getElementById('tab-'+tab).classList.add('active');
          document.getElementById('content-'+tab).classList.add('active');
        }
      </script>
    `;
    }

    function generateTabContent(data, type, allVendors, empid, ts, sig) {
        const vendorFilterOptions = [...new Set(
            data.map(row => `<option value="${row.vendorId}">${row.vendor}</option>`)
        )].join('');
        const itemFilterOptions = [...new Set(
            data.map(row => `<option value="${row.itemId}">${row.item}</option>`)
        )].join('');

        const vendorEditOptions = (allVendors || [])
            .map(v => `<option value="${v.id}">${v.name}</option>`)
            .join('');

        // Buttons differ by tab
        const buttons = (type === 'approval')
            ? `<button type="submit" onclick="return submitUpdates()">Submit Updates</button>
         <button type="submit" onclick="return setAction('${type}','approve')">Approve</button>
         <button type="submit" onclick="return setAction('${type}','reject')">Reject</button>`
            : `<button type="submit" onclick="return setAction1('${type}','createpo')">Create Purchase Order</button>
               <button type="submit" onclick="return setAction1('${type}','scrappo')">Scrape PO</button>`;

        return `
      <form method="POST">
        <input type="hidden" name="custpage_selected_ids" id="custpage_selected_ids_${type}" />
        <input type="hidden" name="custpage_action_type" id="custpage_action_type_${type}" />
        <!-- signed session back -->
        <input type="hidden" name="custpage_empid" value="${String(empid || '').replace(/"/g, '&quot;')}" />
        <input type="hidden" name="custpage_ts"    value="${String(ts || '').replace(/"/g, '&quot;')}" />
        <input type="hidden" name="custpage_sig"   value="${String(sig || '').replace(/"/g, '&quot;')}" />
  
        ${type === 'approval' ? `
          <input type="hidden" name="custpage_changed_ids" id="custpage_changed_ids" />
          <input type="hidden" name="custpage_has_dirty" id="custpage_has_dirty" value="0" />
        ` : ''}
  
        <div class="filter-row">
          <div>
            <label>Filter by Vendor:</label>
            <select id="vendorFilter_${type}" multiple>${vendorFilterOptions}</select>
          </div>
          <div>
            <label>Filter by Item:</label>
            <select id="itemFilter_${type}" multiple>${itemFilterOptions}</select>
          </div>
        </div>
  
        <div>${buttons}</div>
  
        <table>
          <thead>
            <tr>
              <th><input type="checkbox" id="checkAll_${type}" onclick="toggleAll(this,'${type}')" /></th>
              <th>Item</th>
              <th>Item Id</th>
              <th>Display Name</th>
              <th>Purchase Memo</th>
              <th>Vendor</th>
              <th>Ordered Qty</th>
              <th>Month of Stocks</th>
              <th>Vendor Price</th>
              <th>Currency</th>
              <th>Vendor Min Qty</th>
              <th>Not Shipped</th>
              <th>Available</th>
              <th>In Transit</th>
              <th>On Order</th>
              <th>Min Month Qty</th>
            </tr>
          </thead>
          <tbody id="tableBody_${type}">
            ${data.map(row => (type === 'approval'
            ? `
                <tr data-id="${row.id}" data-item="${row.itemId}" data-vendor="${row.vendorId}">
                  <td><input type="checkbox" class="selectLine_${type}" name="selectLine_${type}" /></td>
                  <td>${row.item}</td>
                  <td>${row.itemId}</td>
                  <td>${row.itemName}</td>
                  <td>
                   <input type="text" name="memo_${row.id}"  class="memoInput"
                          value="${(row.memo || '').replace(/"/g, '&quot;')}"
                          data-original="${(row.memo || '').replace(/"/g, '&quot;')}"
                          oninput="markDirty(this)"
                          onchange="markDirty(this)"
                   />
                  </td>
                  <td>
                    <select name="vendor_${row.id}" class="vendSel" data-original="${row.vendorId}">
                      ${vendorEditOptions.replace(`value="${row.vendorId}"`, `value="${row.vendorId}" selected`)}
                    </select>
                  </td>
                  <td>
                    <input type="number" name="qty_${row.id}" class="qtyInput"
                           value="${row.orderQty || ''}" data-original="${row.orderQty || ''}"
                           min="0" step="1" inputmode="numeric" pattern="\\d*"
                           oninput="this.value=this.value.replace(/[^0-9]/g,''); markDirty(this)"
                           onchange="markDirty(this)" />
                  </td>
                  <td>${row.mstock}</td>
                  <td>
                    <input type="number" name="cost_${row.id}" class="costInput"
                           value="${row.cost || ''}" data-original="${row.cost || ''}"
                           min="0" step="0.01" inputmode="decimal"
                           oninput="this.value=this.value.replace(/[^0-9.]/g,'').replace(/(\\..*)\\./g,'$1'); markDirty(this)"
                           onchange="markDirty(this)" />
                  </td>
                  <td>${row.currency}</td>
                  <td>${row.minQty}</td>
                  <td>${row.backOrdered}</td>
                  <td>${row.available}</td>
                  <td>${row.inTransit}</td>
                  <td>${row.onOrder}</td>
                  <td>${row.minMonthQty}</td>
                </tr>
              `
            : `
                <tr data-id="${row.id}" data-item="${row.itemId}" data-vendor="${row.vendorId}">
                  <td><input type="checkbox" class="selectLine_${type}" name="selectLine_${type}" /></td>
                  <td>${row.item}</td><td>${row.itemId}</td><td>${row.itemName}</td><td>${row.memo}</td><td>${row.vendor}</td><td>${row.orderQty}</td><td>${row.mstock}</td><td>${row.cost}</td><td>${row.currency}</td><td>${row.minQty}</td>
                  <td>${row.available}</td><td>${row.inTransit}</td><td>${row.onOrder}</td><td>${row.backOrdered}</td><td>${row.minMonthQty}</td>
                </tr>
              `)
        ).join('')}
          </tbody>
        </table>
  
        <script>
          // ------- Shared helpers -------
          window.getSelectedIds = function(type){
            return Array.from(document.querySelectorAll('#tableBody_'+type+' tr:not([style*="display: none"]) .selectLine_'+type+':checked'))
              .map(cb => cb.closest('tr').dataset.id);
          };
          window.toggleAll = function(src, type){
            const visibleRows = document.querySelectorAll('#tableBody_'+type+' tr:not([style*="display: none"])');
            visibleRows.forEach(row => {
              const cb = row.querySelector('.selectLine_'+type);
              if (cb) cb.checked = src.checked;
            });
          };
          window.filterTable = function(type){
            const vendorValues = Array.from(document.getElementById('vendorFilter_'+type).selectedOptions).map(o=>o.value);
            const itemValues   = Array.from(document.getElementById('itemFilter_'+type).selectedOptions).map(o=>o.value);
            document.querySelectorAll('#tableBody_'+type+' tr').forEach(row => {
              const show = (!vendorValues.length || vendorValues.includes(row.dataset.vendor)) &&
                           (!itemValues.length   || itemValues.includes(row.dataset.item));
              row.style.display = show ? '' : 'none';
            });
          };
          document.addEventListener('DOMContentLoaded', function(){
            document.getElementById('vendorFilter_${type}').addEventListener('change', function(){ filterTable('${type}'); });
            document.getElementById('itemFilter_${type}').addEventListener('change',   function(){ filterTable('${type}'); });
            const checkAllBox = document.getElementById('checkAll_${type}');
            if (checkAllBox){ checkAllBox.addEventListener('change', function(){ toggleAll(this,'${type}'); }); }
          });
  
          ${type === 'approval' ? `
          // ------- Approval-only: dirty tracking + submit updates + guards -------
          function setDirtyFlag(isDirty){
            var f = document.getElementById('custpage_has_dirty');
            if (f) f.value = isDirty ? '1' : '0';
          }
          function hasDirty(){
            return document.querySelector('#tableBody_approval tr.dirty') !== null;
          }
          function markDirty(el){
            const tr = el.closest('tr'); if (!tr) return;
            const vend = tr.querySelector('select.vendSel');
            const qty  = tr.querySelector('input.qtyInput');
            const memo  = tr.querySelector('input.memoInput');
            const cost = tr.querySelector('input.costInput');
            const vendDirty = vend && vend.value !== (vend.getAttribute('data-original') || '');
            const qtyDirty  = qty  && (qty.value  !== (qty.getAttribute('data-original')  || ''));
            const costDirty = cost && (cost.value !== (cost.getAttribute('data-original') || ''));
            const memoDirty = memo && (memo.value !== (memo.getAttribute('data-original') || ''));
            const isDirtyRow = !!(vendDirty || qtyDirty || costDirty || memoDirty);
            tr.classList.toggle('dirty', isDirtyRow);
            setDirtyFlag(hasDirty());
          }
          function submitUpdates(){
            const changed = Array.from(document.querySelectorAll('#tableBody_approval tr.dirty')).map(tr=>tr.dataset.id);
            if (!changed.length){ alert('Please update at least one line before submitting.'); return false; }
            document.getElementById('custpage_changed_ids').value = changed.join(',');
            document.getElementById('custpage_action_type_approval').value = 'save';
            return true;
          }
          window.setAction = function(type, action){
            if (type === 'approval' && hasDirty()){
              alert('You have unsaved changes. Please click "Submit Updates" first, then ' + action + '.');
              return false;
            }
            const selectedIds = getSelectedIds(type);
            if(!selectedIds.length){ alert('Please select at least one line to perform this action.'); return false; }
            document.getElementById('custpage_selected_ids_'+type).value = selectedIds.join(',');
            document.getElementById('custpage_action_type_'+type).value = action;
            return true;
          };
          // init change listener for vendor selects
          document.addEventListener('DOMContentLoaded', function(){
            document.querySelectorAll('#tableBody_approval select.vendSel').forEach(function(sel){
              sel.addEventListener('change', function(){ markDirty(sel); });
            });
          });
          ` : `
          // CreatePO tab keeps the simple setAction1
          window.setAction1 = function(type, action){
            const selectedIds = getSelectedIds(type);
            if(!selectedIds.length){ alert('Please select at least one line to perform this action.'); return false; }
            document.getElementById('custpage_selected_ids_'+type).value = selectedIds.join(',');
            document.getElementById('custpage_action_type_'+type).value = action;
            return true;
          };
          `}
        </script>
      </form>
    `;
    }

    function getAllVendors() {
        var out = [];
        var s = search.create({
            type: search.Type.VENDOR,
            filters: [['isinactive', 'is', 'F']],
            columns: ['internalid', 'altname']
        });
        out.push({ id: '', name: '' });
        s.run().each(function (r) {

            out.push({ id: r.getValue('internalid'), name: r.getValue('altname') });
            return true;
        });
        return out;
    }


    return { onRequest };
});
