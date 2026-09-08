/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 *
 * Vendor Deduction / Remittance Reconciliation Portal
 * ----------------------------------------------------
 * Landing dashboard  : all Bill Credits for a given vendor, grouped by
 *                      remittance number (custbody_note_to_vendor), showing
 *                      credit total / applied / unapplied at a glance.
 * Detail view        : click a remittance -> see the Bill Credit plus every
 *                      Bill (auto-created by the BI deduction automation)
 *                      tied to the same remittance number. Approve applies
 *                      that bill's line on the Bill Credit's apply sublist.
 *                      Claim # / Claim Status are editable inline.
 *
 * ============================================================================
 * CONFIRMED FIELD IDs:
 *   FIELD_REMIT_NO      = custbody_note_to_vendor  (Bill Credit + Bill)
 *   FIELD_CLAIM_NO      = custbody_mi_claim_no     (Bill)
 *   FIELD_CLAIM_STATUS  = custbody_mi_claim_status (Bill) -- List/Record, single value
 *
 *   Claim Status options are loaded dynamically from Custom List internal ID
 *   1416 (record type "customlist1416") at request time, cached 1 hour via
 *   N/cache. Nothing is hardcoded -- if the list changes in NetSuite, this
 *   Suitelet picks it up on the next cache refresh with no code changes.
 * ============================================================================
 */
define(['N/record', 'N/search', 'N/https', 'N/log', 'N/runtime', 'N/url', 'N/cache'],
  function (record, search, https, log, runtime, url, cache) {

    // ---- Field / config constants ------------------------------------------------
    var FIELD_REMIT_NO     = 'custbody_note_to_vendor';
    var FIELD_CLAIM_NO     = 'custbody_mi_claim_no';
    var FIELD_CLAIM_STATUS = 'custbody_mi_claim_status';
    var FIELD_BILL_VALIDATOR = 'custbody_vendbill_validator';
    var FIELD_VALIDATOR_CHECK = 'custbody_vendbill_validator_check';

    var PARAM_FINAL_APPROVER = 'custscript_vdr_final_approver';
    var PARAM_APPROVAL_USERS = 'custscript_vdr_approval_users';

    var APPROVAL_STATUS_PENDING = '1';
    var APPROVAL_STATUS_APPROVED = '2';

    // NEW: Vendor must have this checkbox checked to appear in the portal
    var FIELD_VENDOR_PORTAL = 'custentity_inclding_deduction_portal';

    var CLAIM_STATUS_LIST_RECTYPE = 'customlist_claim_status';
    var CLAIM_STATUS_CACHE_NAME   = 'mi_claim_status_options_v1';
    var CLAIM_STATUS_CACHE_TTL    = 60 * 60;

    var APPROVE_STATUS_NAME = 'Approved - Ready to Apply';

    // =========================================================================
    // Entry point
    // =========================================================================
    function onRequest(context) {
      try {
        if (context.request.method === 'GET') {
          var action = context.request.parameters.action;

          if (action === 'detail') {
            renderDetail(context);
          } else if (action === 'billapproval') {
            renderBillApproval(context);
          } else {
            renderDashboard(context);
          }

        } else {
          handlePost(context);
        }

      } catch (e) {
        log.error('onRequest error', e);

        if (context.request.method === 'POST') {
          context.response.write(JSON.stringify({
            success: false,
            error: e.message
          }));

        } else {
          context.response.write(
            renderShell(
              'Error',
              '<div class="err-box">' + esc(e.message) + '</div>'
            )
          );
        }
      }
    }

    // =========================================================================
    // POST handler -- AJAX actions from the detail page
    // =========================================================================
    function handlePost(context) {
      context.response.setHeader({
        name: 'Content-Type',
        value: 'application/json'
      });

      var body;

      try {
        body = JSON.parse(context.request.body);
      } catch (e) {
        context.response.write(JSON.stringify({
          success: false,
          error: 'Bad request body'
        }));
        return;
      }

      if (body.action === 'approve') {
        var result = applyBillCredit(body.creditId, body.billId);
        context.response.write(JSON.stringify(result));
        return;
      }

      if (body.action === 'updateClaim') {
        var result2 = updateClaimFields(
          body.billId,
          body.claimNo,
          body.claimStatus
        );

        context.response.write(JSON.stringify(result2));
        return;
      }

      if (body.action === 'approveBill') {
        var result3 = approvePendingBill(body.billId);
        context.response.write(JSON.stringify(result3));
        return;
      }

      context.response.write(JSON.stringify({
        success: false,
        error: 'Unknown action'
      }));
    }

    // Apply one bill's line on the Bill Credit's apply sublist, full amount.
    function applyBillCredit(creditId, billId) {
      try {

        var creditRec = record.load({
          type: record.Type.VENDOR_CREDIT,
          id: creditId,
          isDynamic: true
        });

        var lineCount = creditRec.getLineCount({
          sublistId: 'apply'
        });

        var foundLine = -1;

        for (var i = 0; i < lineCount; i++) {

          var docId =
            creditRec.getSublistValue({
              sublistId: 'apply',
              fieldId: 'internalid',
              line: i
            }) ||
            creditRec.getSublistValue({
              sublistId: 'apply',
              fieldId: 'doc',
              line: i
            });

          if (String(docId) === String(billId)) {
            foundLine = i;
            break;
          }
        }

        if (foundLine === -1) {
          return {
            success: false,
            error:
              'This bill does not appear on the Bill Credit\'s apply sublist. ' +
              'It may be a different subsidiary/currency, or not open. Check manually in NetSuite.'
          };
        }

        creditRec.selectLine({
          sublistId: 'apply',
          line: foundLine
        });

        creditRec.setCurrentSublistValue({
          sublistId: 'apply',
          fieldId: 'apply',
          value: true
        });

        creditRec.commitLine({
          sublistId: 'apply'
        });

        creditRec.save();

        var approveStatusId =
          findClaimStatusIdByName(APPROVE_STATUS_NAME);

        if (approveStatusId) {

          var values = {};
          values[FIELD_CLAIM_STATUS] = approveStatusId;

          record.submitFields({
            type: record.Type.VENDOR_BILL,
            id: billId,
            values: values
          });

        } else {

          log.audit(
            'applyBillCredit',
            'Could not find claim status list value named "' +
            APPROVE_STATUS_NAME +
            '" in ' +
            CLAIM_STATUS_LIST_RECTYPE +
            ' -- skipped auto status update.'
          );
        }

        return {
          success: true
        };

      } catch (e) {

        log.error('applyBillCredit error', e);

        return {
          success: false,
          error: e.message
        };
      }
    }

    function updateClaimFields(billId, claimNo, claimStatus) {
      try {

        var values = {};

        if (claimNo !== undefined) {
          values[FIELD_CLAIM_NO] = claimNo;
        }

        if (
          claimStatus !== undefined &&
          claimStatus !== ''
        ) {
          values[FIELD_CLAIM_STATUS] = claimStatus;
        }

        record.submitFields({
          type: record.Type.VENDOR_BILL,
          id: billId,
          values: values
        });

        return {
          success: true
        };

      } catch (e) {

        log.error('updateClaimFields error', e);

        return {
          success: false,
          error: e.message
        };
      }
    }

    // =========================================================================
    // Two-stage Vendor Bill approval
    // =========================================================================
    function getApprovalConfig() {

      var script = runtime.getCurrentScript();
      var finalApprover = script.getParameter({
        name: PARAM_FINAL_APPROVER
      });
      var approvalUsers = script.getParameter({
        name: PARAM_APPROVAL_USERS
      });

      return {
        finalApproverId:
          finalApprover === null || finalApprover === undefined
            ? ''
            : String(finalApprover),

        overviewUserMap: parseEmployeeIdList(approvalUsers)
      };
    }

    function parseEmployeeIdList(value) {

      var out = {};

      String(value || '')
        .split(/[\s,;]+/)
        .forEach(function (id) {
          var cleanId = String(id || '').trim();

          if (cleanId) {
            out[cleanId] = true;
          }
        });

      return out;
    }

    function getCurrentUserId() {

      var currentUser = runtime.getCurrentUser();

      return currentUser && currentUser.id !== undefined
        ? String(currentUser.id)
        : '';
    }

    function approvePendingBill(billId) {
      try {

        if (!billId) {
          throw new Error('Bill is required.');
        }

        var currentUserId = getCurrentUserId();
        var config = getApprovalConfig();
        var billRec = record.load({
          type: record.Type.VENDOR_BILL,
          id: billId,
          isDynamic: false
        });

        var approvalStatus = String(
          billRec.getValue({
            fieldId: 'approvalstatus'
          }) || ''
        );

        var vendorId = String(
          billRec.getValue({
            fieldId: 'entity'
          }) || ''
        );

        var validatorId = String(
          billRec.getValue({
            fieldId: FIELD_BILL_VALIDATOR
          }) || ''
        );

        var validatorChecked = !!billRec.getValue({
          fieldId: FIELD_VALIDATOR_CHECK
        });

        if (!vendorId || !isVendorEnabledForPortal(vendorId)) {
          throw new Error(
            'This bill is not available in the Deduction Portal.'
          );
        }

        if (approvalStatus !== APPROVAL_STATUS_PENDING) {
          throw new Error('This bill is no longer pending approval.');
        }

        if (!validatorChecked) {

          if (!validatorId || currentUserId !== validatorId) {
            throw new Error(
              'Only the validator assigned to this bill can complete validator approval.'
            );
          }

          var validatorValues = {};
          validatorValues[FIELD_VALIDATOR_CHECK] = true;

          record.submitFields({
            type: record.Type.VENDOR_BILL,
            id: billId,
            values: validatorValues,
            options: {
              enableSourcing: false,
              ignoreMandatoryFields: false
            }
          });

          return {
            success: true,
            stage: 'validator',
            message: 'Validator approval completed.'
          };
        }

        if (
          !config.finalApproverId ||
          currentUserId !== config.finalApproverId
        ) {
          throw new Error(
            'Only the configured final approver can approve this bill.'
          );
        }

        record.submitFields({
          type: record.Type.VENDOR_BILL,
          id: billId,
          values: {
            approvalstatus: Number(APPROVAL_STATUS_APPROVED)
          },
          options: {
            enableSourcing: false,
            ignoreMandatoryFields: false
          }
        });

        return {
          success: true,
          stage: 'final',
          message: 'Bill approved.'
        };

      } catch (e) {

        log.error('approvePendingBill error', e);

        return {
          success: false,
          error: e.message
        };
      }
    }

    function isVendorEnabledForPortal(vendorId) {

      var lookup = search.lookupFields({
        type: search.Type.VENDOR,
        id: vendorId,
        columns: [FIELD_VENDOR_PORTAL]
      });
      var value = lookup[FIELD_VENDOR_PORTAL];

      return value === true || value === 'T';
    }

    // =========================================================================
    // Claim Status list -- loaded dynamically, cached
    // =========================================================================
    function getClaimStatusOptions() {

      var c = cache.getCache({
        name: CLAIM_STATUS_CACHE_NAME,
        scope: cache.Scope.PROTECTED
      });

      var jsonStr = c.get({
        key: 'options',
        loader: loadClaimStatusOptionsFromList,
        ttl: CLAIM_STATUS_CACHE_TTL
      });

      return JSON.parse(jsonStr);
    }

    function loadClaimStatusOptionsFromList() {

      var out = [];

      var s = search.create({
        type: CLAIM_STATUS_LIST_RECTYPE,

        columns: [
          search.createColumn({
            name: 'internalid'
          }),

          search.createColumn({
            name: 'name',
            sort: search.Sort.ASC
          })
        ]
      });

      s.run().each(function (r) {

        out.push({
          id: r.getValue({
            name: 'internalid'
          }),

          name: r.getValue({
            name: 'name'
          })
        });

        return true;
      });

      log.audit(
        'loadClaimStatusOptionsFromList',
        'Loaded ' +
        out.length +
        ' values from ' +
        CLAIM_STATUS_LIST_RECTYPE
      );

      return JSON.stringify(out);
    }

    function findClaimStatusIdByName(name) {

      var opts = getClaimStatusOptions();
      var lower = String(name).toLowerCase();

      for (var i = 0; i < opts.length; i++) {

        if (
          String(opts[i].name).toLowerCase() === lower
        ) {
          return opts[i].id;
        }
      }

      return null;
    }

    // =========================================================================
    // NEW: Vendors enabled for Deduction Portal
    // =========================================================================
    function getEligibleVendors() {

      var out = [];

      search.create({
        type: search.Type.VENDOR,

        filters: [
          ['isinactive', 'is', 'F'],
          'AND',
          [FIELD_VENDOR_PORTAL, 'is', 'T']
        ],

        columns: [
          search.createColumn({
            name: 'companyname',
            sort: search.Sort.ASC
          }),

          search.createColumn({
            name: 'entityid'
          }),

          search.createColumn({
            name: 'internalid'
          })
        ]

      }).run().each(function (r) {

        var vendorId = String(
          r.getValue({
            name: 'internalid'
          })
        );

        var vendorName =
          r.getValue({
            name: 'companyname'
          }) ||
          r.getValue({
            name: 'entityid'
          }) ||
          ('Vendor #' + vendorId);

        out.push({
          id: vendorId,
          name: vendorName
        });

        return true;
      });

      return out;
    }

    // =========================================================================
    // Dashboard
    // =========================================================================
    function renderDashboard(context) {

      // NEW: Vendor list comes from checkbox instead of hardcoded vendor
      var vendors = getEligibleVendors();

      var vendorMap = {};

      vendors.forEach(function (v) {
        vendorMap[String(v.id)] = v.name;
      });

      // Blank means All Vendors
      var vendorId =
        context.request.parameters.vendorid || '';

      // Prevent URL from manually passing non-enabled vendor
      if (
        vendorId &&
        !vendorMap[String(vendorId)]
      ) {
        vendorId = '';
      }

      var vendorIds;

      if (vendorId) {

        vendorIds = [String(vendorId)];

      } else {

        vendorIds = vendors.map(function (v) {
          return String(v.id);
        });
      }

      var credits =
        vendorIds.length > 0
          ? getVendorCredits(vendorIds, vendorMap)
          : [];

      var totalCredits = credits.length;
      var totalAmt = 0;
      var appliedAmt = 0;
      var unappliedAmt = 0;
      var openCount = 0;

      credits.forEach(function (c) {

        totalAmt += c.total;
        appliedAmt += c.applied;
        unappliedAmt += c.unapplied;

        if (c.unapplied > 0.001) {
          openCount++;
        }
      });

      // -----------------------------------------------------------------------
      // NEW: Vendor Filter
      // -----------------------------------------------------------------------
      var vendorOptions =
        '<option value="' +
        esc(selfUrl({})) +
        '"' +
        (!vendorId ? ' selected' : '') +
        '>All Vendors</option>';

      vendors.forEach(function (v) {

        var selected =
          String(v.id) === String(vendorId)
            ? ' selected'
            : '';

        vendorOptions +=
          '<option value="' +
          esc(
            selfUrl({
              vendorid: v.id
            })
          ) +
          '"' +
          selected +
          '>' +
          esc(v.name) +
          '</option>';
      });

      var filterHtml =
        '<div class="filter-bar">' +
          '<label for="vendor-filter">Vendor</label>' +
          '<select id="vendor-filter" onchange="window.location.href=this.value">' +
            vendorOptions +
          '</select>' +
        '</div>';

      var rows = credits.map(function (c, idx) {

        var statusBadge =
          c.unapplied <= 0.001
            ? '<span class="badge badge-green">Fully Applied</span>'
            : (
                c.applied > 0
                  ? '<span class="badge badge-amber">Partially Applied</span>'
                  : '<span class="badge badge-blue">Open</span>'
              );

        return (
          '<tr class="dr" data-idx="' +
          idx +
          '">' +

          // NEW: Vendor Name
          '<td>' +
          esc(c.vendorName) +
          '</td>' +

          '<td>' +
          esc(c.remitNo || '(none)') +
          '</td>' +

          '<td><a href="' +
          creditUrl(c.id) +
          '" target="_blank">' +
          esc(c.tranid) +
          '</a></td>' +

          '<td>' +
          esc(c.trandate) +
          '</td>' +

          '<td class="num">' +
          fmtMoney(c.total) +
          '</td>' +

          '<td class="num">' +
          fmtMoney(c.applied) +
          '</td>' +

          '<td class="num">' +
          fmtMoney(c.unapplied) +
          '</td>' +

          '<td class="num">' +
          c.billCount +
          '</td>' +

          '<td>' +
          statusBadge +
          '</td>' +

          '<td><a class="row-link" href="' +
          selfUrl({
            action: 'detail',
            vendorid: c.vendorId,
            creditid: c.id,
            remitno: c.remitNo || '',
            dashboardvendorid: vendorId
          }) +
          '">Open &rsaquo;</a></td>' +

          '<td><a class="row-link" href="' +
          selfUrl({
            action: 'billapproval',
            vendorid: c.vendorId,
            dashboardvendorid: vendorId
          }) +
          '">Approve Bills &rsaquo;</a></td>' +

          '</tr>'
        );
      }).join('');

      var html =

        filterHtml +

        tiles([
          {
            label: 'Open Bill Credits',
            value: totalCredits,
            color: '#003764'
          },

          {
            label: 'Total Credit Amount',
            value: fmtMoney(totalAmt),
            color: '#003764'
          },

          {
            label: 'Applied to Date',
            value: fmtMoney(appliedAmt),
            color: '#3D7A41'
          },

          {
            label: 'Unapplied (Pending)',
            value: fmtMoney(unappliedAmt),
            color: '#B95C00'
          }
        ]) +

        '<div class="card">' +

        '<table class="tbl">' +

        '<thead><tr>' +

        // NEW
        '<th>Vendor</th>' +

        '<th>Remittance #</th>' +

        '<th>Bill Credit</th>' +

        '<th>Date</th>' +

        '<th class="num">Credit Total</th>' +

        '<th class="num">Applied</th>' +

        '<th class="num">Unapplied</th>' +

        '<th class="num">Bills</th>' +

        '<th>Status</th>' +

        '<th></th>' +

        '<th>Approve Bills</th>' +

        '</tr></thead>' +

        '<tbody id="tbl-body">' +

        (
          rows ||
          '<tr><td colspan="11" class="empty">' +
          'No open Bill Credits found.' +
          '</td></tr>'
        ) +

        '</tbody>' +

        '</table>' +

        '</div>';

      var selectedVendorName =
        vendorId && vendorMap[vendorId]
          ? vendorMap[vendorId]
          : 'All Vendors';

      context.response.write(
        renderShell(
          'Deduction & Remittance Portal',
          html,
          'Vendor: ' + esc(selectedVendorName)
        )
      );
    }

    // =========================================================================
    // Pending Vendor Bills -- validator approval, then final approval
    // =========================================================================
    function renderBillApproval(context) {

      var vendorId = String(
        context.request.parameters.vendorid || ''
      );
      var dashboardVendorId = String(
        context.request.parameters.dashboardvendorid || ''
      );
      var eligibleVendors = getEligibleVendors();
      var vendorMap = {};

      eligibleVendors.forEach(function (v) {
        vendorMap[String(v.id)] = v.name;
      });

      if (!vendorId) {
        throw new Error('Vendor is required.');
      }

      if (!vendorMap[vendorId]) {
        throw new Error(
          'This vendor is not enabled for the Deduction Portal.'
        );
      }

      var currentUserId = getCurrentUserId();
      var config = getApprovalConfig();
      var canViewAll = !!config.overviewUserMap[currentUserId];
      var allPendingBills = getPendingApprovalBills(vendorId);

      var bills = allPendingBills.filter(function (b) {

        if (canViewAll) {
          return true;
        }

        if (
          !b.validatorApproved &&
          b.validatorId === currentUserId
        ) {
          return true;
        }

        return (
          b.validatorApproved &&
          !!config.finalApproverId &&
          config.finalApproverId === currentUserId
        );
      });

      var rows = bills.map(function (b) {

        var isValidatorAction =
          !b.validatorApproved &&
          !!b.validatorId &&
          b.validatorId === currentUserId;

        var isFinalAction =
          b.validatorApproved &&
          !!config.finalApproverId &&
          config.finalApproverId === currentUserId;

        var statusBadge = b.validatorApproved
          ? '<span class="badge badge-amber">Final Approval</span>'
          : '<span class="badge badge-blue">Validator Approval</span>';

        var actionHtml =
          isValidatorAction || isFinalAction
            ? '<button class="btn-bill-approve">Approve</button>'
            : '<span class="muted">&mdash;</span>';

        return (
          '<tr data-billid="' + esc(b.id) + '">' +

          '<td>' + esc(b.vendorName) + '</td>' +

          '<td><a href="' + billUrl(b.id) + '" target="_blank">' +
          esc(b.tranid) +
          '</a></td>' +

          '<td>' + esc(b.trandate) + '</td>' +

          '<td class="num">' + fmtMoney(b.amount) + '</td>' +

          '<td>' + esc(b.validatorName || '(not assigned)') + '</td>' +

          '<td>' +
          (
            b.validatorApproved
              ? '<span class="badge badge-green">Yes</span>'
              : '<span class="badge badge-blue">No</span>'
          ) +
          '</td>' +

          '<td>' + statusBadge + '</td>' +

          '<td>' + actionHtml + '</td>' +

          '<td class="row-msg"></td>' +

          '</tr>'
        );
      }).join('');

      var backUrl = dashboardVendorId
        ? selfUrl({
            vendorid: dashboardVendorId
          })
        : selfUrl({});

      var html =
        '<div style="margin-bottom:14px">' +
          '<a href="' + backUrl + '" style="color:#36677D;font-size:12px">' +
          '&lsaquo; Back to Dashboard' +
          '</a>' +
        '</div>' +

        tiles([
          {
            label: 'Pending Bills',
            value: bills.length,
            color: '#003764'
          },
          {
            label: 'Validator Approval',
            value: bills.filter(function (b) {
              return !b.validatorApproved;
            }).length,
            color: '#36677D'
          },
          {
            label: 'Final Approval',
            value: bills.filter(function (b) {
              return b.validatorApproved;
            }).length,
            color: '#B95C00'
          }
        ]) +

        '<div class="card">' +
          '<table class="tbl">' +
            '<thead><tr>' +
              '<th>Vendor</th>' +
              '<th>Bill #</th>' +
              '<th>Date</th>' +
              '<th class="num">Amount</th>' +
              '<th>Validator</th>' +
              '<th>Validator Approved?</th>' +
              '<th>Approval Status</th>' +
              '<th>Action</th>' +
              '<th></th>' +
            '</tr></thead>' +
            '<tbody id="approval-tbl-body">' +
            (
              rows ||
              '<tr><td colspan="9" class="empty">' +
              'No pending bills are available for this user.' +
              '</td></tr>'
            ) +
            '</tbody>' +
          '</table>' +
        '</div>' +

        billApprovalClientScript();

      context.response.write(
        renderShell(
          'Approve Bills - ' + vendorMap[vendorId],
          html,
          'Logged in as employee #' + currentUserId
        )
      );
    }

    // =========================================================================
    // Detail -- one remittance: Bill Credit summary + all related Bills
    // =========================================================================
    function renderDetail(context) {

      var vendorId =
        context.request.parameters.vendorid || '';

      var creditId =
        context.request.parameters.creditid;

      var remitNo =
        context.request.parameters.remitno || '';

      // NEW: Used only to return user to same dashboard filter
      var dashboardVendorId =
        context.request.parameters.dashboardvendorid || '';

      // NEW: Verify vendor is enabled for portal
      var eligibleVendors = getEligibleVendors();
      var vendorMap = {};

      eligibleVendors.forEach(function (v) {
        vendorMap[String(v.id)] = v.name;
      });

      if (!vendorId) {
        throw new Error('Vendor is required.');
      }

      if (!vendorMap[String(vendorId)]) {
        throw new Error(
          'This vendor is not enabled for the Deduction Portal.'
        );
      }

      var vendorName =
        vendorMap[String(vendorId)];

      var creditInfo =
        getCreditWithApplyLines(creditId);

      var bills =
        getBillsForRemit(
          vendorId,
          remitNo,
          creditInfo.applyLines
        );

      var claimStatusOptions =
        getClaimStatusOptions();

      var billRowsJson =
        JSON.stringify(bills)
          .replace(/<\/script>/gi, '<\\/script>');

      var rows = bills.map(function (b, idx) {

        var claimStatusOpts =
          claimStatusOptions.map(function (opt) {

            var isSelected =
              String(opt.id) ===
              String(b.claimStatusId);

            return (
              '<option value="' +
              esc(opt.id) +
              '"' +
              (isSelected ? ' selected' : '') +
              '>' +
              esc(opt.name) +
              '</option>'
            );

          }).join('');

        var approveDisabled =
          b.applied ? 'disabled' : '';

        var approveLabel =
          b.applied ? 'Applied' : 'Approve';

        return (
          '<tr data-idx="' +
          idx +
          '" data-billid="' +
          b.id +
          '">' +

          // NEW: Vendor Name
          '<td>' +
          esc(vendorName) +
          '</td>' +

          '<td><a href="' +
          billUrl(b.id) +
          '" target="_blank">' +
          esc(b.tranid) +
          '</a></td>' +

          '<td>' +
          esc(b.trandate) +
          '</td>' +

          '<td class="num">' +
          fmtMoney(b.amount) +
          '</td>' +

          '<td><input type="text" class="claimno-input" value="' +
          esc(b.claimNo || '') +
          '" placeholder="Claim #" /></td>' +

          '<td><select class="claimstatus-select">' +
          claimStatusOpts +
          '</select></td>' +

          '<td>' +
          (
            b.applied
              ? '<span class="badge badge-green">Applied</span>'
              : '<span class="badge badge-blue">Not Applied</span>'
          ) +
          '</td>' +

          '<td>' +
            '<button class="btn-approve" ' +
            approveDisabled +
            '>' +
            approveLabel +
            '</button> ' +

            '<button class="btn-save-claim">' +
            'Save Claim Info' +
            '</button>' +
          '</td>' +

          '<td class="row-msg"></td>' +

          '</tr>'
        );
      }).join('');

      // NEW: Return to same vendor filter selected on dashboard
      var backUrl =
        dashboardVendorId
          ? selfUrl({
              vendorid: dashboardVendorId
            })
          : selfUrl({});

      var html =

        '<div style="margin-bottom:14px">' +

          '<a href="' +
          backUrl +
          '" style="color:#36677D;font-size:12px">' +
          '&lsaquo; Back to Dashboard' +
          '</a>' +

        '</div>' +

        tiles([
          {
            label: 'Remittance #',
            value: esc(remitNo || '(none)'),
            color: '#003764'
          },

          {
            label: 'Credit Total',
            value: fmtMoney(creditInfo.total),
            color: '#003764'
          },

          {
            label: 'Applied',
            value: fmtMoney(creditInfo.applied),
            color: '#3D7A41'
          },

          {
            label: 'Unapplied',
            value: fmtMoney(creditInfo.unapplied),
            color: '#B95C00'
          }
        ]) +

        '<div class="card">' +

        '<div class="card-hdr">' +

        'Bill Credit ' +

        '<a href="' +
        creditUrl(creditInfo.id) +
        '" target="_blank">' +
        esc(creditInfo.tranid) +
        '</a> — Related Bills' +

        '</div>' +

        '<table class="tbl">' +

        '<thead><tr>' +

        // NEW
        '<th>Vendor</th>' +

        '<th>Bill #</th>' +

        '<th>Date</th>' +

        '<th class="num">Amount</th>' +

        '<th>Claim #</th>' +

        '<th>Claim Status</th>' +

        '<th>Applied?</th>' +

        '<th>Actions</th>' +

        '<th></th>' +

        '</tr></thead>' +

        '<tbody id="tbl-body">' +

        (
          rows ||
          '<tr><td colspan="9" class="empty">' +
          'No bills found tied to this remittance number.' +
          '</td></tr>'
        ) +

        '</tbody>' +

        '</table>' +

        '</div>' +

        '<script type="application/json" id="row-data">' +
        billRowsJson +
        '</script>' +

        '<script type="application/json" id="ctx-data">' +
        JSON.stringify({
          creditId: creditInfo.id
        }) +
        '</script>' +

        clientScript();

      context.response.write(
        renderShell(
          'Remittance ' +
          esc(remitNo) +
          ' - ' +
          esc(vendorName),

          html,

          'Vendor: ' +
          esc(vendorName) +
          ' (#' +
          esc(vendorId) +
          ')'
        )
      );
    }

    // =========================================================================
    // Data access
    // =========================================================================
    function getVendorName(vendorId) {
      try {

        var lookup = search.lookupFields({
          type: search.Type.VENDOR,
          id: vendorId,
          columns: [
            'companyname',
            'entityid'
          ]
        });

        return (
          lookup.companyname ||
          lookup.entityid ||
          ('Vendor #' + vendorId)
        );

      } catch (e) {

        return 'Vendor #' + vendorId;
      }
    }

    function getPendingApprovalBills(vendorId) {

      var out = [];

      search.create({
        type: search.Type.VENDOR_BILL,

        filters: [
          ['mainline', 'is', 'T'],
          'AND',
          ['entity', 'anyof', vendorId],
          'AND',
          ['approvalstatus', 'anyof', APPROVAL_STATUS_PENDING]
        ],

        columns: [
          search.createColumn({
            name: 'internalid'
          }),

          search.createColumn({
            name: 'entity'
          }),

          search.createColumn({
            name: 'tranid'
          }),

          search.createColumn({
            name: 'trandate',
            sort: search.Sort.ASC
          }),

          search.createColumn({
            name: 'amount'
          }),

          search.createColumn({
            name: FIELD_BILL_VALIDATOR
          }),

          search.createColumn({
            name: FIELD_VALIDATOR_CHECK
          })
        ]

      }).run().each(function (r) {

        var validatorApproved = r.getValue({
          name: FIELD_VALIDATOR_CHECK
        });

        out.push({
          id: String(
            r.getValue({
              name: 'internalid'
            })
          ),

          vendorId: String(
            r.getValue({
              name: 'entity'
            })
          ),

          vendorName:
            r.getText({
              name: 'entity'
            }) || getVendorName(vendorId),

          tranid:
            r.getValue({
              name: 'tranid'
            }),

          trandate:
            r.getValue({
              name: 'trandate'
            }),

          amount: Math.abs(
            parseFloat(
              r.getValue({
                name: 'amount'
              })
            ) || 0
          ),

          validatorId: String(
            r.getValue({
              name: FIELD_BILL_VALIDATOR
            }) || ''
          ),

          validatorName:
            r.getText({
              name: FIELD_BILL_VALIDATOR
            }) || '',

          validatorApproved:
            validatorApproved === true ||
            validatorApproved === 'T'
        });

        return true;
      });

      return out;
    }

    // =========================================================================
    // UPDATED ONLY FOR MULTIPLE VENDORS
    // =========================================================================
    function getVendorCredits(vendorIds, vendorMap) {

      var out = [];

      var s = search.create({
        type: search.Type.VENDOR_CREDIT,

        filters: [
          ['mainline', 'is', 'T'],
          'AND',
          ['entity', 'anyof', vendorIds]
        ],

        columns: [
          search.createColumn({
            name: 'internalid'
          }),

          // NEW
          search.createColumn({
            name: 'entity'
          }),

          search.createColumn({
            name: 'tranid'
          }),

          search.createColumn({
            name: 'trandate'
          }),

          search.createColumn({
            name: 'total'
          }),

          search.createColumn({
            name: 'amountremaining'
          }),

          search.createColumn({
            name: FIELD_REMIT_NO
          })
        ]
      });

      var billCounts =
        getBillCountsByRemit(vendorIds);

      s.run().each(function (r) {

        var id =
          r.getValue({
            name: 'internalid'
          });

        // NEW
        var vendorId = String(
          r.getValue({
            name: 'entity'
          })
        );

        // NEW
        var vendorName =
          vendorMap[vendorId] ||
          r.getText({
            name: 'entity'
          }) ||
          ('Vendor #' + vendorId);

        var total =
          Math.abs(
            parseFloat(
              r.getValue({
                name: 'total'
              })
            ) || 0
          );

        var unapplied =
          Math.abs(
            parseFloat(
              r.getValue({
                name: 'amountremaining'
              })
            ) || 0
          );

        if (unapplied <= 0.001) {
          return true;
        }

        var remitNo =
          r.getValue({
            name: FIELD_REMIT_NO
          }) || '';

        // NEW: vendor + remittance key
        var countKey =
          vendorId +
          '||' +
          remitNo;

        out.push({

          id: id,

          // NEW
          vendorId: vendorId,

          // NEW
          vendorName: vendorName,

          tranid:
            r.getValue({
              name: 'tranid'
            }),

          trandate:
            r.getValue({
              name: 'trandate'
            }),

          remitNo: remitNo,

          total: total,

          applied:
            total - unapplied,

          unapplied:
            unapplied,

          billCount:
            billCounts[countKey] || 0
        });

        return true;
      });

      return out;
    }

    // =========================================================================
    // UPDATED ONLY FOR MULTIPLE VENDORS
    // =========================================================================
    function getBillCountsByRemit(vendorIds) {

      var counts = {};

      var s = search.create({
        type: search.Type.VENDOR_BILL,

        filters: [
          ['mainline', 'is', 'T'],
          'AND',
          ['entity', 'anyof', vendorIds]
        ],

        columns: [

          // NEW: group by vendor as well
          search.createColumn({
            name: 'entity',
            summary: search.Summary.GROUP
          }),

          search.createColumn({
            name: FIELD_REMIT_NO,
            summary: search.Summary.GROUP
          }),

          search.createColumn({
            name: 'internalid',
            summary: search.Summary.COUNT
          })
        ]
      });

      s.run().each(function (r) {

        var vendorId = String(
          r.getValue({
            name: 'entity',
            summary: search.Summary.GROUP
          })
        );

        var key =
          r.getValue({
            name: FIELD_REMIT_NO,
            summary: search.Summary.GROUP
          }) || '';

        var cnt =
          parseInt(
            r.getValue({
              name: 'internalid',
              summary: search.Summary.COUNT
            }),
            10
          ) || 0;

        // NEW: Prevent same remit number across different vendors mixing counts
        counts[
          vendorId +
          '||' +
          key
        ] = cnt;

        return true;
      });

      return counts;
    }

    // Loads a single Bill Credit's apply sublist -> [{docId, apply, amount}]
    function getApplyLinesForCredit(creditId) {

      var lines = [];

      try {

        var rec = record.load({
          type: record.Type.VENDOR_CREDIT,
          id: creditId,
          isDynamic: false
        });

        var lineCount =
          rec.getLineCount({
            sublistId: 'apply'
          });

        for (var i = 0; i < lineCount; i++) {

          var isApplying =
            rec.getSublistValue({
              sublistId: 'apply',
              fieldId: 'apply',
              line: i
            });

          var docId =
            rec.getSublistValue({
              sublistId: 'apply',
              fieldId: 'internalid',
              line: i
            }) ||
            rec.getSublistValue({
              sublistId: 'apply',
              fieldId: 'doc',
              line: i
            });

          var amount =
            parseFloat(
              rec.getSublistValue({
                sublistId: 'apply',
                fieldId: 'amount',
                line: i
              })
            ) || 0;

          lines.push({
            docId: docId,
            apply: !!isApplying,
            amount: amount
          });
        }

      } catch (e) {

        log.error(
          'getApplyLinesForCredit error for credit ' +
          creditId,
          e
        );
      }

      return lines;
    }

    function getCreditWithApplyLines(creditId) {

      var lookup = search.lookupFields({
        type: search.Type.VENDOR_CREDIT,
        id: creditId,
        columns: [
          'tranid',
          'trandate',
          'total',
          FIELD_REMIT_NO
        ]
      });

      var applyLines =
        getApplyLinesForCredit(creditId);

      var total =
        Math.abs(
          parseFloat(lookup.total) || 0
        );

      var appliedSum =
        applyLines.reduce(function (sum, l) {
          return sum + (
            l.apply
              ? l.amount
              : 0
          );
        }, 0);

      return {
        id: creditId,
        tranid: lookup.tranid,
        trandate: lookup.trandate,
        total: total,
        applied: appliedSum,
        unapplied: total - appliedSum,
        applyLines: applyLines
      };
    }

    function getBillsForRemit(
      vendorId,
      remitNo,
      applyLines
    ) {

      var appliedMap = {};

      applyLines.forEach(function (l) {
        appliedMap[l.docId] = l.apply;
      });

      var out = [];

      var s = search.create({
        type: search.Type.VENDOR_BILL,

        filters: [
          ['mainline', 'is', 'T'],
          'AND',
          ['entity', 'anyof', vendorId],
          'AND',
          [FIELD_REMIT_NO, 'is', remitNo]
        ],

        columns: [
          search.createColumn({
            name: 'internalid'
          }),

          search.createColumn({
            name: 'tranid'
          }),

          search.createColumn({
            name: 'trandate'
          }),

          search.createColumn({
            name: 'amount'
          }),

          search.createColumn({
            name: FIELD_CLAIM_NO
          }),

          search.createColumn({
            name: FIELD_CLAIM_STATUS
          })
        ]
      });

      s.run().each(function (r) {

        var id =
          r.getValue({
            name: 'internalid'
          });

        out.push({

          id: id,

          tranid:
            r.getValue({
              name: 'tranid'
            }),

          trandate:
            r.getValue({
              name: 'trandate'
            }),

          amount:
            Math.abs(
              parseFloat(
                r.getValue({
                  name: 'amount'
                })
              ) || 0
            ),

          claimNo:
            r.getValue({
              name: FIELD_CLAIM_NO
            }),

          claimStatusId:
            r.getValue({
              name: FIELD_CLAIM_STATUS
            }) || '',

          claimStatusText:
            r.getText({
              name: FIELD_CLAIM_STATUS
            }) || 'Pending Review',

          applied:
            !!appliedMap[id]
        });

        return true;
      });

      return out;
    }

    // =========================================================================
    // Rendering helpers
    // =========================================================================
    function tiles(items) {

      return (
        '<div class="tiles">' +

        items.map(function (t) {

          return (
            '<div class="tile">' +

            '<div class="tile-lbl">' +
            esc(t.label) +
            '</div>' +

            '<div class="tile-val" style="color:' +
            t.color +
            '">' +
            t.value +
            '</div>' +

            '</div>'
          );

        }).join('') +

        '</div>'
      );
    }

    function fmtMoney(n) {

      n = n || 0;

      var neg =
        n < 0;

      var s =
        Math.abs(n).toLocaleString(
          'en-US',
          {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
          }
        );

      return (
        (neg ? '-$' : '$') +
        s
      );
    }

    function esc(v) {

      if (
        v === null ||
        v === undefined
      ) {
        return '';
      }

      return String(v)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    function selfUrl(params) {

      var script =
        runtime.getCurrentScript();

      return url.resolveScript({
        scriptId: script.id,
        deploymentId: script.deploymentId,
        params: params || {}
      });
    }

    function creditUrl(id) {
      return (
        '/app/accounting/transactions/vendcred.nl?id=' +
        id
      );
    }

    function billUrl(id) {
      return (
        '/app/accounting/transactions/vendbill.nl?id=' +
        id
      );
    }

    function renderShell(
      title,
      bodyHtml,
      subtitle
    ) {

      return (
        '<!DOCTYPE html>' +

        '<html>' +

        '<head>' +

        '<meta charset="utf-8">' +

        '<title>' +
        esc(title) +
        '</title>' +

        '<style>' +
        baseCss() +
        '</style>' +

        '</head>' +

        '<body>' +

        '<div class="page">' +

        '<div class="hdr">' +

        '<h1>' +
        esc(title) +
        '</h1>' +

        (
          subtitle
            ? '<p>' +
              esc(subtitle) +
              '</p>'
            : ''
        ) +

        '</div>' +

        bodyHtml +

        '</div>' +

        '</body>' +

        '</html>'
      );
    }

    function baseCss() {

      return [

        'body{font-family:Arial,Helvetica,sans-serif;background:#F5F5F5;margin:0;color:#1a1a1a;}',

        '.page{padding:24px;}',

        '.hdr{background:#003764;border-radius:8px;padding:18px 22px;margin-bottom:20px;}',

        '.hdr h1{color:#fff;font-size:19px;font-weight:600;margin:0;}',

        '.hdr p{color:#B8D4E8;font-size:12px;margin:4px 0 0;}',

        // NEW: Vendor filter style
        '.filter-bar{background:#fff;border:1px solid #D9D9D9;border-radius:8px;padding:12px 16px;margin-bottom:18px;}',

        '.filter-bar label{font-size:11px;font-weight:600;color:#6B6B6B;text-transform:uppercase;margin-right:10px;}',

        '.filter-bar select{min-width:250px;padding:6px 8px;border:1px solid #D9D9D9;border-radius:4px;font-size:12px;}',

        '.tiles{display:flex;gap:14px;margin-bottom:18px;flex-wrap:wrap;}',

        '.tile{background:#fff;border:1px solid #D9D9D9;border-radius:8px;padding:14px 18px;min-width:150px;}',

        '.tile-lbl{font-size:11px;color:#6B6B6B;text-transform:uppercase;letter-spacing:.4px;margin-bottom:6px;}',

        '.tile-val{font-size:22px;font-weight:700;}',

        '.card{background:#fff;border:1px solid #D9D9D9;border-radius:8px;padding:18px;}',

        '.card-hdr{font-size:13px;font-weight:600;color:#003764;margin-bottom:12px;}',

        '.tbl{width:100%;border-collapse:collapse;font-size:12px;}',

        '.tbl th{text-align:left;color:#6B6B6B;font-size:10.5px;text-transform:uppercase;letter-spacing:.3px;border-bottom:1px solid #D9D9D9;padding:8px 8px;}',

        '.tbl td{padding:8px;border-bottom:1px solid #F0F0F0;vertical-align:middle;}',

        '.tbl .num{text-align:right;}',

        '.empty{text-align:center;color:#999;padding:24px;}',

        'a{color:#36677D;text-decoration:none;} a:hover{text-decoration:underline;}',

        '.row-link{font-weight:600;}',

        '.badge{display:inline-block;padding:2px 9px;border-radius:10px;font-size:10.5px;font-weight:600;}',

        '.badge-green{background:#E6F2E6;color:#3D7A41;}',

        '.badge-amber{background:#FBEEDC;color:#B95C00;}',

        '.badge-blue{background:#E5EEF3;color:#36677D;}',

        '.claimno-input{width:110px;padding:4px 6px;font-size:12px;border:1px solid #D9D9D9;border-radius:4px;}',

        '.claimstatus-select{padding:4px 6px;font-size:12px;border:1px solid #D9D9D9;border-radius:4px;}',

        'button{padding:5px 10px;font-size:11px;border-radius:4px;border:1px solid #D9D9D9;background:#fff;cursor:pointer;}',

        '.btn-approve{background:#3D7A41;color:#fff;border-color:#3D7A41;}',

        '.btn-bill-approve{background:#3D7A41;color:#fff;border-color:#3D7A41;}',

        '.btn-approve[disabled]{background:#D9D9D9;border-color:#D9D9D9;color:#888;cursor:not-allowed;}',

        '.btn-bill-approve[disabled]{background:#D9D9D9;border-color:#D9D9D9;color:#888;cursor:not-allowed;}',

        '.btn-save-claim{color:#36677D;}',

        '.row-msg{font-size:11px;color:#3D7A41;}',

        '.muted{color:#999;}',

        '.err-box{background:#FBEAEA;color:#B00020;border:1px solid #E8B4B4;border-radius:6px;padding:14px;}'

      ].join('');
    }

    function billApprovalClientScript() {

      return (
        '<script>' +

        '(function(){' +

        'var body = document.getElementById("approval-tbl-body");' +

        'if(!body) return;' +

        'function post(payload, cb){' +

          'var xhr = new XMLHttpRequest();' +

          'xhr.open("POST", window.location.pathname + window.location.search);' +

          'xhr.setRequestHeader("Content-Type","application/json");' +

          'xhr.onload = function(){' +

            'var res = {};' +

            'try{' +
              'res = JSON.parse(xhr.responseText);' +
            '}catch(e){}' +

            'cb(res);' +

          '};' +

          'xhr.onerror = function(){' +
            'cb({success:false,error:"Network error"});' +
          '};' +

          'xhr.send(JSON.stringify(payload));' +

        '}' +

        'body.addEventListener("click", function(e){' +

          'if(!e.target.classList.contains("btn-bill-approve")) return;' +

          'var tr = e.target.closest("tr[data-billid]");' +

          'if(!tr) return;' +

          'var button = e.target;' +

          'var msgCell = tr.querySelector(".row-msg");' +

          'button.disabled = true;' +

          'button.textContent = "Approving...";' +

          'post({' +
            'action:"approveBill",' +
            'billId:tr.getAttribute("data-billid")' +
          '}, function(res){' +

            'if(res.success){' +
              'window.location.reload();' +
              'return;' +
            '}' +

            'button.disabled = false;' +

            'button.textContent = "Approve";' +

            'msgCell.style.color = "#B00020";' +

            'msgCell.textContent = res.error || "Approval failed";' +

          '});' +

        '});' +

        '})();' +

        '</script>'
      );
    }

    // =========================================================================
    // Client-side JS for detail page
    // =========================================================================
    function clientScript() {

      return (
        '<script>' +

        '(function(){' +

        'var ROWS = JSON.parse(document.getElementById("row-data").textContent);' +

        'var CTX = JSON.parse(document.getElementById("ctx-data").textContent);' +

        'var body = document.getElementById("tbl-body");' +

        'function post(payload, cb){' +

          'var xhr = new XMLHttpRequest();' +

          'xhr.open("POST", window.location.pathname + window.location.search);' +

          'xhr.setRequestHeader("Content-Type","application/json");' +

          'xhr.onload = function(){' +

            'var res = {};' +

            'try{' +
              'res = JSON.parse(xhr.responseText);' +
            '}catch(e){}' +

            'cb(res);' +

          '};' +

          'xhr.send(JSON.stringify(payload));' +

        '}' +

        'body.addEventListener("click", function(e){' +

          'var tr = e.target.closest("tr[data-billid]");' +

          'if(!tr) return;' +

          'var billId = tr.getAttribute("data-billid");' +

          'var msgCell = tr.querySelector(".row-msg");' +

          'if(e.target.classList.contains("btn-approve")){' +

            'e.target.disabled = true;' +

            'e.target.textContent = "Applying...";' +

            'post({' +
              'action:"approve",' +
              'creditId:CTX.creditId,' +
              'billId:billId' +
            '}, function(res){' +

              'if(res.success){' +

                'e.target.textContent = "Applied";' +

                // Vendor column added, therefore Applied column moved from 5 to 6
                'var badge = tr.children[6];' +

                'badge.innerHTML = "<span class=\\"badge badge-green\\">Applied</span>";' +

                'msgCell.textContent = "Applied successfully";' +

              '}else{' +

                'e.target.disabled = false;' +

                'e.target.textContent = "Approve";' +

                'msgCell.style.color = "#B00020";' +

                'msgCell.textContent = res.error || "Failed to apply";' +

              '}' +

            '});' +

          '}' +

          'if(e.target.classList.contains("btn-save-claim")){' +

            'var claimNo = tr.querySelector(".claimno-input").value;' +

            'var claimStatus = tr.querySelector(".claimstatus-select").value;' +

            'e.target.disabled = true;' +

            'e.target.textContent = "Saving...";' +

            'post({' +

              'action:"updateClaim",' +

              'billId:billId,' +

              'claimNo:claimNo,' +

              'claimStatus:claimStatus' +

            '}, function(res){' +

              'e.target.disabled = false;' +

              'e.target.textContent = "Save Claim Info";' +

              'msgCell.style.color = res.success ? "#3D7A41" : "#B00020";' +

              'msgCell.textContent = res.success ? "Saved" : (res.error || "Save failed");' +

            '});' +

          '}' +

        '});' +

        '})();' +

        '</script>'
      );
    }

    return {
      onRequest: onRequest
    };

  });
