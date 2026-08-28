/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 */
define(['N/record', 'N/log', 'N/ui/serverWidget'], function(record, log, serverWidget) {

    const VENDOR_BILL_RECORD_TYPE = 'vendorbill';
    const APPROVAL_STATUS_PENDING_APPROVAL = '1';
    const APPROVAL_STATUS_APPROVED = '2';
    const LCL = '11437';
    const METRO = '442';

    function beforeLoad(context) {
        try {
            if (
                context.type !== context.UserEventType.VIEW &&
                context.type !== context.UserEventType.EDIT
            ) {
                return;
            }

            const bill = context.newRecord;

            if (!bill.getValue('custbody_created_from_email_capture')) return;
            const venId = bill.getValue('entity');
            if (venId !== LCL && venId !== METRO) return;

            if (!bill || bill.type !== VENDOR_BILL_RECORD_TYPE) {
                return;
            }

            const badgeConfig = getApprovalStatusBadgeConfig(bill);

            if (!badgeConfig) {
                return;
            }

            addApprovalStatusBadge(context.form, badgeConfig);
        } catch (error) {
            log.error({
                title: 'Molisana bill approval status beforeLoad failed',
                details: error
            });
        }
    }

    function getApprovalStatusBadgeConfig(bill) {
        const approvalStatusId = String(bill.getValue({ fieldId: 'approvalstatus' }) || '');
        const validatorCheck = isChecked(bill.getValue({ fieldId: 'custbody_vendbill_validator_check' }));

        if (approvalStatusId === APPROVAL_STATUS_PENDING_APPROVAL && !validatorCheck) {
            return {
                label: 'Validator Approval',
                theme: 'warning',
                title: 'Pending validator approval'
            };
        }

        if (approvalStatusId === APPROVAL_STATUS_PENDING_APPROVAL && validatorCheck) {
            return {
                label: 'Final Approval',
                theme: 'warning',
                title: 'Pending final approval'
            };
        }

        if (approvalStatusId === APPROVAL_STATUS_APPROVED) {
            return {
                label: 'Approved',
                theme: 'success',
                title: 'Approved'
            };
        }

        return null;
    }

    function isChecked(value) {
        return value === true || value === 'T' || value === 'true' || value === '1';
    }

    function addApprovalStatusBadge(form, badgeConfig) {
        const htmlField = form.addField({
            id: 'custpage_molisana_bill_approval_status_html',
            label: 'Molisana Bill Approval Status',
            type: serverWidget.FieldType.INLINEHTML
        });

        htmlField.defaultValue = `
            <style>
                #custpage_molisana_bill_approval_status_html_fs,
                #custpage_molisana_bill_approval_status_html_fs_lbl {
                    display: none !important;
                }

                #custpage_molisana_bill_approval_status_badge {
                    display: inline-flex !important;
                    align-items: center;
                    margin-left: 8px;
                    padding: 2px 8px;
                    border: 1px solid #b8c6d6;
                    border-radius: 3px;
                    background: #eef2f6;
                    color: #34495e;
                    font-size: 12px !important;
                    font-weight: 600;
                    line-height: 16px;
                    text-transform: uppercase;
                    vertical-align: middle;
                    white-space: nowrap;
                }

                #custpage_molisana_bill_approval_status_badge[data-badge-theme="success"] {
                    border-color: #8fb88f;
                    background: #e8f4e8;
                    color: #246b24;
                }

                #custpage_molisana_bill_approval_status_badge[data-badge-theme="warning"] {
                    border-color: #d8b35a;
                    background: #fff4d8;
                    color: #7a5200;
                }
            </style>
            <script>
                (function () {
                    var badgeLabel = ${toJavaScriptString(badgeConfig.label)};
                    var badgeTheme = ${toJavaScriptString(badgeConfig.theme)};
                    var badgeTitle = ${toJavaScriptString(badgeConfig.title)};
                    var badgeId = 'custpage_molisana_bill_approval_status_badge';

                    function findRecordTitle() {
                        return document.querySelector('.uir-record-name') ||
                            document.querySelector('.uir-page-title-firstline h1') ||
                            document.querySelector('.uir-page-title h1') ||
                            document.querySelector('#div__title h1') ||
                            document.querySelector('h1');
                    }

                    function addBadge() {
                        var title = findRecordTitle();

                        if (!title || document.getElementById(badgeId)) {
                            return Boolean(title);
                        }

                        var badge = document.createElement('span');
                        badge.id = badgeId;
                        badge.className = 'uir-record-status';
                        badge.textContent = badgeLabel;
                        badge.setAttribute('data-badge-theme', badgeTheme);
                        badge.setAttribute('title', badgeTitle);
                        badge.setAttribute('aria-label', 'Molisana bill approval status: ' + badgeLabel);

                        title.appendChild(document.createTextNode(' '));
                        title.appendChild(badge);

                        return true;
                    }

                    if (document.readyState === 'loading') {
                        document.addEventListener('DOMContentLoaded', addBadge);
                    } else {
                        addBadge();
                    }

                    var attempts = 0;
                    var timer = window.setInterval(function () {
                        attempts += 1;

                        if (addBadge() || attempts >= 20) {
                            window.clearInterval(timer);
                        }
                    }, 250);
                }());
            </script>
        `;
    }

    function toJavaScriptString(value) {
        return JSON.stringify(String(value || '')).replace(/<\//g, '<\\/');
    }
  

    function afterSubmit(context) {
        if (context.type !== context.UserEventType.CREATE) return;

      try {

        var newRecord = context.newRecord;

        // Get the header field value
        var headerInboundShipment = newRecord.getValue({ fieldId: 'custbody_vendbill_related_to_inbship' });
        var venRefNum = newRecord.getValue({fieldId: 'tranid'})

        if (!headerInboundShipment) {
            log.debug('No Header Value', 'No value found for custbody_vendbill_related_to_inbship');
            return;
        }

        // Load the record in edit mode
        var vendorBill = record.load({
            type: 'vendorbill',
            id: newRecord.id,
            isDynamic: false
        });

        // Get the number of lines in the item sublist
        var lineCount = vendorBill.getLineCount({ sublistId: 'item' });

        // Loop through the lines in reverse to safely remove lines
        for (var i = lineCount - 1; i >= 0; i--) {
            var lineInboundShipment = vendorBill.getSublistValue({
                sublistId: 'item',
                fieldId: 'custcol_mi_related_inbound',
                line: i
            });

            var venRef = vendorBill.getSublistValue({
                sublistId: 'item',
                fieldId: 'custcol_mi_vendor_ref_number',
                line: i
            });

            // If line's field does not match the header field, remove the line
            if (lineInboundShipment != headerInboundShipment || venRef != venRefNum) {
                vendorBill.removeLine({
                    sublistId: 'item',
                    line: i,
                    ignoreRecalc: true
                });
                log.debug('Line Removed', 'Line ' + i + ' removed due to mismatch.');
            }
        }

        // Save the updated record
        vendorBill.save({ ignoreMandatoryFields: true });

        log.debug('Vendor Bill Updated', 'Non-matching lines removed successfully.');
        
        
      } catch (error) {
        log.error('error',error)
      }
    }

    return {
        beforeLoad: beforeLoad,
        afterSubmit: afterSubmit
    };
});
