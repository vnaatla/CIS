import { Access } from '../core/auth/access'
import type { IconName } from '../shared/icon.component'

/**
 * How the item appears in the VB6 desktop (MDIMain.frm). Everything is migrated ("Everything" scope), these flags are kept for traceability.
 * - hidden: designer Visible = 0 · disabled: designer Enabled = 0 · never-opened: no menu item / handler commented out, the form is still in
 *   the project · nothing: menu item without a click handler and without a form (nothing to build, listed only for completeness).
 */
export type DesktopState = 'visible' | 'hidden' | 'disabled' | 'never-opened' | 'nothing'

export interface MenuItem {
  /** Desktop caption. */
  label: string
  /** Web route (absolute). */
  path: string
  /** USRAccess position (CheckAccess). */
  access: number
  /** Legacy VB6 form(s) the screen replaces. */
  form: string
  /** Sub-menu heading in the desktop (shown as a group in the sidebar). */
  group?: string
  desktop?: DesktopState
  /** Migration phase of the plan (docs/menu-mapping.md). */
  phase: number
  /** True once the screen is built; until then the sidebar shows it disabled ("not migrated yet"). */
  migrated?: boolean
}

export interface MenuModule {
  key: string
  /** Desktop top-menu caption. */
  label: string
  icon: IconName
  /** Module overview route. */
  path: string
  items: MenuItem[]
}

export const MENU: MenuModule[] = [
  {
    key: 'membership',
    label: 'Membership',
    icon: 'users',
    path: '/membership',
    items: [
      { label: 'MBM Adjustment', path: '/membership/adjustment', access: Access.MbmAdjustment, form: 'FrmMBMCLNAdj', phase: 2, migrated: true },
      { label: 'MBM Enquiry', path: '/membership/enquiry', access: Access.MbmEnquiry, form: 'FrmMBMCLNEnq, FrmMBMAccount', phase: 2 },
      { label: 'MBM Topup Adj', path: '/membership/topup', access: Access.MbmTopup, form: 'FrmMBMTopupLmt', phase: 2, migrated: true },
      { label: 'MBM Tracking', path: '/membership/tracking', access: Access.MbmTracking, form: 'FrmNewMBMTracking', group: 'Tracking', phase: 10 },
      { label: 'MBM Tracking (Simple)', path: '/membership/tracking-simple', access: Access.MbmTracking, form: 'frmNewMBMTrackingSimple', group: 'Tracking', phase: 10 },
      { label: 'MBM Tracking (Jupiter)', path: '/membership/tracking-jupiter', access: Access.MbmTrackingJupiter, form: 'FrmNewMBMTracking2', group: 'Tracking', phase: 10 },
      { label: 'MBM Tracking (Jupiter - Simple)', path: '/membership/tracking-jupiter-simple', access: Access.MbmTrackingJupiter, form: 'frmNewMBMTrackingRepSimple', group: 'Tracking', phase: 10 },
      { label: 'MBM Tracking (old)', path: '/membership/tracking-legacy', access: Access.MbmTracking, form: 'FrmMBMTracking', group: 'Tracking', desktop: 'never-opened', phase: 10 },
    ],
  },
  {
    key: 'mco-fees',
    label: 'MCO Fees',
    icon: 'receipt',
    path: '/mco-fees',
    items: [
      { label: 'MCO Invoice', path: '/mco-fees/invoice', access: Access.McoInvoice, form: 'frmMCOInvoiceNew', phase: 5 },
      { label: 'MCO Invoice GST', path: '/mco-fees/invoice-gst', access: Access.McoInvoice, form: 'frmMCOFeeGST', phase: 5 },
      { label: 'MCO Invoice GST Zeroised', path: '/mco-fees/invoice-gst-zeroised', access: Access.McoInvoice, form: 'frmMCOFeeGSTZeroised', phase: 5 },
      { label: 'Fee Listing', path: '/mco-fees/fee-listing', access: Access.McoFeeListing, form: 'frmMCOFeeListNew', phase: 5 },
      { label: 'Refund Listing', path: '/mco-fees/refund-listing', access: Access.McoFeeListing, form: 'frmMCORefundListing', phase: 5 },
    ],
  },
  {
    key: 'invoice',
    label: 'Invoice',
    icon: 'file',
    path: '/invoice',
    items: [
      { label: 'Invoice Control', path: '/invoice/control', access: Access.InvoiceClinic, form: '—', desktop: 'nothing', phase: 3 },
      { label: 'Clinic', path: '/invoice/clinic', access: Access.InvoiceClinic, form: 'frmInvoice, frmWorksheet', group: 'Clinic', phase: 3 },
      { label: 'Member', path: '/invoice/member', access: Access.InvoiceMember, form: 'frmInvoiceMBM', group: 'Clinic', phase: 3 },
      { label: 'Adjustment', path: '/invoice/adjustment', access: Access.InvoiceAdjustment, form: 'frmInvoiceAdj', group: 'Clinic', phase: 3 },
      { label: 'Adjustment (After Posting)', path: '/invoice/adjustment-after-posting', access: Access.InvoiceAdjustment, form: 'frmInvoiceAdjAfterPost', group: 'Clinic', phase: 3 },
      { label: 'Enquiry', path: '/invoice/enquiry', access: Access.InvoiceAdjustment, form: 'FrmInvoiceEnq', group: 'Clinic', phase: 3 },
      { label: 'Batch Transfer', path: '/invoice/batch-transfer', access: Access.InvoiceAdjustment, form: 'FrmInvoiceBatchNoUpd', group: 'Clinic', phase: 3 },
      { label: 'Adjustment (After Bordx)', path: '/invoice/adjustment-after-bordx', access: Access.InvoiceAdjustmentAfterBordx, form: 'frmInvoiceAdjAftBordx', group: 'Clinic', phase: 3 },
      { label: 'Email Notification (Payment)', path: '/invoice/email/payment', access: Access.EmailNotification, form: 'frmEmailNotification', group: 'Email Notification', phase: 10 },
      { label: 'Email Notification (Providers)', path: '/invoice/email/providers', access: Access.EmailNotification, form: 'frmEmailNotificationProvider', group: 'Email Notification', phase: 10 },
      { label: 'Email Notification (Rejection)', path: '/invoice/email/rejection', access: Access.EmailNotification, form: 'frmEmailRejectedClaim', group: 'Email Notification', phase: 10 },
      { label: 'General Exclusions', path: '/invoice/general-exclusions', access: Access.GeneralExclusions, form: 'frmGeneralBenefitExclusion', group: 'General Exclusions', phase: 2 },
      { label: 'General Exclusions (Enquiry)', path: '/invoice/general-exclusions/enquiry', access: Access.GeneralExclusionsEnquiry, form: 'frmGeneralBenefitExcEnq', group: 'General Exclusions', phase: 2, migrated: true },
      { label: 'Claim ASCII (Bintulu)', path: '/invoice/claim-ascii/bintulu', access: Access.ClaimAscii, form: 'frmClaimDetailsASCII', group: 'Claim ASCII', phase: 10 },
      { label: 'Claim ASCII (Toshiba)', path: '/invoice/claim-ascii/toshiba', access: Access.ClaimAscii, form: 'frmClaimASCIIToshiba', group: 'Claim ASCII', phase: 10 },
      { label: 'Claim ASCII (Singer)', path: '/invoice/claim-ascii/singer', access: Access.ClaimAscii, form: 'frmClaimASCIISinger', group: 'Claim ASCII', phase: 10 },
      { label: 'Claim ASCII (Utusan Melayu)', path: '/invoice/claim-ascii/utusan', access: Access.ClaimAscii, form: 'frmClaimASCIIUtusan', group: 'Claim ASCII', phase: 10 },
      { label: 'Claim ASCII (OMNI)', path: '/invoice/claim-ascii/omni', access: Access.ClaimAscii, form: 'frmClaimDataOMNI', group: 'Claim ASCII', phase: 10 },
      { label: 'Claim ASCII (TIMS)', path: '/invoice/claim-ascii/tims', access: Access.ClaimAscii, form: 'frmClaimASCIITIMS', group: 'Claim ASCII', phase: 10 },
      { label: 'Claim ASCII (CIMB)', path: '/invoice/claim-ascii/cimb', access: Access.ClaimAscii, form: 'frmClaimASCIICIMB', group: 'Claim ASCII', phase: 10 },
      { label: 'Med Providers (OMNI)', path: '/invoice/claim-ascii/omni-providers', access: Access.ClaimAscii, form: 'frmMedProviderOMNI', group: 'Claim ASCII', desktop: 'never-opened', phase: 10 },
      { label: 'Claims Data', path: '/invoice/claim-ascii/claims-data', access: Access.ClaimAscii, form: 'frmClaimsdata', group: 'Claim ASCII', desktop: 'never-opened', phase: 10 },
      { label: 'Claims Excel (Sin Chew)', path: '/invoice/claim-ascii/sinchew', access: Access.ClaimAscii, form: 'frmClaimsExcelSinChew', group: 'Claim ASCII', desktop: 'never-opened', phase: 10 },
      { label: 'Dental', path: '/invoice/dental', access: Access.InvoiceClinic, form: '— (no form in the project)', group: 'Dental', desktop: 'nothing', phase: 11 },
      { label: 'Dental Member', path: '/invoice/dental-member', access: Access.InvoiceClinic, form: '— (no form in the project)', group: 'Dental', desktop: 'nothing', phase: 11 },
      { label: 'Dental Adjustment', path: '/invoice/dental-adjustment', access: Access.InvoiceClinic, form: '— (no form in the project)', group: 'Dental', desktop: 'nothing', phase: 11 },
      { label: 'Invoice Update', path: '/invoice/update', access: Access.InvoiceAdjustment, form: 'frmUpdateCLNInvoice', group: 'Other', desktop: 'hidden', phase: 3 },
      { label: 'OP Batch/Invoice', path: '/invoice/op-batch', access: Access.InvoiceAdjustment, form: 'frmUpdateInvoiceBatch', group: 'Other', phase: 3 },
      { label: 'Case Transfer', path: '/invoice/case-transfer', access: Access.CaseTransfer, form: 'frmMembershipTransfer', group: 'Other', phase: 2, migrated: true },
      { label: 'Tracking', path: '/invoice/tracking', access: Access.InvoiceTracking, form: 'FrmInvoiceTracking', group: 'Other', phase: 10 },
      { label: 'Tracking (Jupiter)', path: '/invoice/tracking-jupiter', access: Access.InvoiceTrackingJupiter, form: 'FrmInvoiceTracking2', group: 'Other', phase: 10 },
      { label: 'MMA/CPT', path: '/invoice/mma-cpt', access: Access.MmaCpt, form: 'frmMMASchedule, frmShowInvestigation', group: 'Other', phase: 3 },
      { label: 'Worksheet Registration', path: '/invoice/worksheet-registration', access: Access.WorksheetRegistration, form: 'frmNewWSReg', group: 'Other', desktop: 'never-opened', phase: 3 },
      { label: 'Worksheet Adjustment', path: '/invoice/worksheet-adjustment', access: Access.WorksheetAdjustment, form: 'frmNewWSAjd', group: 'Other', desktop: 'never-opened', phase: 3 },
    ],
  },
  {
    key: 'claim',
    label: 'Claim',
    icon: 'shield',
    path: '/claim',
    items: [{ label: 'Pass to Claim', path: '/claim/pass-to-claim', access: Access.PassToClaim, form: 'frmPass2Claim', phase: 3 }],
  },
  {
    key: 'claim-bordx',
    label: 'Claim Bordx',
    icon: 'list',
    path: '/claim-bordx',
    items: [
      { label: 'Insert Worksheet', path: '/claim-bordx/insert-worksheet', access: Access.BordxInsert, form: 'frmWSBordx', group: 'Worksheet', phase: 6 },
      { label: 'Insert Invoice', path: '/claim-bordx/insert-invoice', access: Access.BordxInsert, form: 'frmWSBordx2', group: 'Worksheet', desktop: 'disabled', phase: 6 },
      { label: 'Remove Worksheet', path: '/claim-bordx/remove-worksheet', access: Access.BordxRemove, form: 'FrmBordxRevise', group: 'Worksheet', phase: 6 },
      { label: 'Update Claims Data', path: '/claim-bordx/update-claims-data', access: Access.BordxPrint, form: 'frmUpdateClaimsData', group: 'Bordx', phase: 3 },
      { label: 'Print Bordx (Crystal)', path: '/claim-bordx/print/crystal', access: Access.BordxPrint, form: 'frmPrintBordx', group: 'Print Bordx', desktop: 'hidden', phase: 6 },
      { label: 'Print Bordx', path: '/claim-bordx/print/standard', access: Access.BordxPrint, form: 'frmPrintBordxExcel', group: 'Print Bordx', phase: 6 },
      { label: 'Print Bordx (SEB/INDAH WATER)', path: '/claim-bordx/print/seb', access: Access.BordxPrint, form: 'frmPrintBordxIBM', group: 'Print Bordx', phase: 6 },
      { label: 'Print Bordx (Prasarana)', path: '/claim-bordx/print/prasarana', access: Access.BordxPrint, form: 'frmPrintBordxSPNB', group: 'Print Bordx', phase: 6 },
      { label: 'Print Bordx (AXA Affin)', path: '/claim-bordx/print/axa-affin', access: Access.BordxPrint, form: 'frmPrintBordxAXAAffin', group: 'Print Bordx', phase: 6 },
      { label: 'Print Bordx (Time Dotcom)', path: '/claim-bordx/print/time-dotcom', access: Access.BordxPrint, form: 'frmPrintBordxTimeDotcom', group: 'Print Bordx', phase: 6 },
      { label: 'Print Bordx (BSN)', path: '/claim-bordx/print/bsn', access: Access.BordxPrint, form: 'frmPrintBordxBSN', group: 'Print Bordx', phase: 6 },
      { label: 'BSN-PrintSummary', path: '/claim-bordx/print/bsn-summary', access: Access.BordxPrint, form: 'frmBSNPrintSummary', group: 'Print Bordx', desktop: 'never-opened', phase: 6 },
      { label: 'Print Bordx (Berjaya Sompo)', path: '/claim-bordx/print/berjaya-sompo', access: Access.BordxPrint, form: 'frmPrintBordxBSI', group: 'Print Bordx', phase: 6 },
      { label: 'Send to Grp. Company', path: '/claim-bordx/delivery/send-to-payor', access: Access.BordxSendToPayor, form: 'FrmSendtoPayor', group: 'Bordx Delivery', phase: 6 },
      { label: 'Returned Fr Grp. Company', path: '/claim-bordx/delivery/returned', access: Access.BordxReturnedFromPayor, form: 'FrmReturnPayor', group: 'Bordx Delivery', phase: 6 },
      { label: 'Send to A/C', path: '/claim-bordx/delivery/send-to-ac', access: Access.BordxSendToAc, form: 'FrmSendtoAC', group: 'Bordx Delivery', phase: 6 },
      { label: 'Send to MNRB', path: '/claim-bordx/delivery/send-to-mnrb', access: Access.BordxSendToPayor, form: 'FrmSendtoMNRB', group: 'Bordx Delivery', desktop: 'never-opened', phase: 6 },
      { label: 'Bor Acc Period', path: '/claim-bordx/ac-period', access: Access.BordxAcPeriod, form: 'FrmACPeriod', group: 'Bordx', phase: 6 },
      { label: 'Bordx Control', path: '/claim-bordx/monitoring/control', access: Access.BordxControl, form: 'FrmBordxControl', group: 'Monitoring', phase: 6 },
      { label: 'Outstanding Bordx', path: '/claim-bordx/monitoring/outstanding', access: Access.BordxOutstanding, form: 'FrmOSBordx', group: 'Monitoring', phase: 6 },
      { label: 'Bordx & Worksheet', path: '/claim-bordx/monitoring/bordx-worksheet', access: Access.BordxWorksheet, form: 'FrmBordxWks', group: 'Monitoring', phase: 6 },
      { label: 'Claim Tracking', path: '/claim-bordx/claim-tracking/standard', access: Access.ClaimTracking, form: 'frmClaimRpt', group: 'Claim Tracking', phase: 10 },
      { label: 'Claim Tracking (Simple)', path: '/claim-bordx/claim-tracking/simple', access: Access.ClaimTracking, form: 'frmClaimRptSimple', group: 'Claim Tracking', phase: 10 },
      { label: 'Claim Tracking 2', path: '/claim-bordx/claim-tracking/v2', access: Access.ClaimTracking2, form: 'FrmClaimRpt2', group: 'Claim Tracking', phase: 10 },
      { label: 'Claim Tracking (Jupiter)', path: '/claim-bordx/claim-tracking/jupiter', access: Access.ClaimTrackingJupiter, form: 'frmNewClaimRpt', group: 'Claim Tracking', phase: 10 },
      { label: 'Claim Tracking (Jupiter - Simple)', path: '/claim-bordx/claim-tracking/jupiter-simple', access: Access.ClaimTrackingJupiter, form: 'frmNewClaimRptSimpleJup', group: 'Claim Tracking', phase: 10 },
      { label: 'Claim Tracking 2 (Jupiter)', path: '/claim-bordx/claim-tracking/v2-jupiter', access: Access.ClaimTracking2Jupiter, form: 'FrmNewClaimRpt2', group: 'Claim Tracking', phase: 10 },
      { label: 'Claim Tracking (SUN System)', path: '/claim-bordx/claim-tracking/sun', access: Access.ClaimTracking, form: 'frmClaimRptSUN', group: 'Claim Tracking', phase: 10 },
      { label: 'ClaimBordxDate', path: '/claim-bordx/bordx-date', access: Access.BordxInsert, form: 'FrmBordxDate', group: 'Bordx', phase: 6 },
    ],
  },
  {
    key: 'lg-audit',
    label: 'LG & Audit',
    icon: 'check',
    path: '/lg-audit',
    items: [
      { label: 'Generate LG', path: '/lg-audit/generate-lg', access: Access.GenerateLg, form: 'frmLGGeneration', group: 'LG', phase: 4 },
      { label: 'Audit review', path: '/lg-audit/audit-review', access: Access.AuditReview, form: 'frmAuditInvoice', group: 'Audit review', phase: 4 },
      { label: 'Audit Tracking', path: '/lg-audit/audit-tracking', access: Access.AuditReview, form: 'FrmAuditInvoiceTracking', group: 'Audit review', phase: 4 },
    ],
  },
  {
    key: 'payment',
    label: 'Payment',
    icon: 'wallet',
    path: '/payment',
    items: [
      { label: 'To Clinic', path: '/payment/clinic', access: Access.PaymentClinic, form: 'frmPaymentClinic', phase: 7 },
      { label: 'To Clinic - Invoice', path: '/payment/clinic-invoice', access: Access.PaymentClinicInvoice, form: 'frmPaymentCLNInvoice, frmPrintChequeRange, frmPrintPVRange, frmInvoiceDetPopup', phase: 7 },
      { label: 'To Member', path: '/payment/member', access: Access.PaymentMember, form: 'frmPaymentMemberAuto08', phase: 7 },
      { label: 'To Member (old)', path: '/payment/member-legacy', access: Access.PaymentMember, form: 'frmPaymentMember', desktop: 'never-opened', phase: 7 },
      { label: 'To Member - Batch', path: '/payment/member-batch', access: Access.PaymentMember, form: '—', desktop: 'nothing', phase: 7 },
      { label: 'Tracking', path: '/payment/tracking', access: Access.PaymentTracking, form: 'FrmPayment', phase: 7 },
      { label: 'Tracking (Jupiter)', path: '/payment/tracking-jupiter', access: Access.PaymentTrackingJupiter, form: 'frmPayment2', phase: 7 },
      { label: 'FFS Payment', path: '/payment/ffs', access: Access.PaymentFfs, form: 'FrmPaymentFFS', phase: 7 },
      { label: 'Pay Member Post', path: '/payment/post/member', access: Access.AcExport, form: 'frmPayMBMPost', desktop: 'never-opened', phase: 7 },
      { label: 'Pay Clinic Post', path: '/payment/post/clinic', access: Access.AcExport, form: 'frmPayClinicPost', desktop: 'never-opened', phase: 7 },
    ],
  },
  {
    key: 'receipt',
    label: 'Receipt',
    icon: 'inbox',
    path: '/receipt',
    items: [
      { label: 'Fr Member', path: '/receipt/member', access: Access.ReceiptMember, form: 'FrmReceiptMember', phase: 8 },
      { label: 'Fr Grp Company', path: '/receipt/payor', access: Access.ReceiptPayor, form: 'FrmReceiptPayor', phase: 8 },
      { label: 'Fr FFS', path: '/receipt/ffs', access: Access.ReceiptFfs, form: 'FrmReceiptFFS', phase: 8 },
      { label: 'Fr MNRB', path: '/receipt/mnrb', access: Access.ReceiptPayor, form: 'FrmReceiptMNRB', desktop: 'never-opened', phase: 8 },
      { label: 'Upload Receipt', path: '/receipt/upload', access: Access.ReceiptPayor, form: 'frmReceiptUpload', phase: 8 },
      { label: 'Upload Receipt (Multiple Receipts)', path: '/receipt/upload-multiple', access: Access.ReceiptPayor, form: 'frmReceiptUpMultiple', phase: 8 },
    ],
  },
  {
    key: 'ac-export',
    label: 'A/C Export',
    icon: 'send',
    path: '/ac-export',
    items: [
      { label: 'Clinic', path: '/ac-export/invoice/clinic', access: Access.AcExport, form: 'frmPostingInvoiceCln', group: 'Invoice', phase: 9 },
      { label: 'Clinic (GST)', path: '/ac-export/invoice/clinic-gst', access: Access.AcExport, form: 'frmPostingInvoiceClnGST', group: 'Invoice', phase: 9 },
      { label: 'Clinic (GST-2017)', path: '/ac-export/invoice/clinic-gst-2017', access: Access.AcExport, form: 'frmPostingInvClnGST2017', group: 'Invoice', phase: 9 },
      { label: 'Clinic (GST-2017 Zeroised)', path: '/ac-export/invoice/clinic-gst-2017-zeroised', access: Access.AcExport, form: 'frmPostingInvClnGST2017Zeroised', group: 'Invoice', desktop: 'never-opened', phase: 9 },
      { label: 'Member', path: '/ac-export/invoice/member', access: Access.AcExport, form: 'frmPostingInvoiceMember', group: 'Invoice', phase: 9 },
      { label: 'Invoice Posting', path: '/ac-export/invoice/posting', access: Access.AcExport, form: 'frmInvPosting', group: 'Invoice', desktop: 'never-opened', phase: 9 },
      { label: 'Clinic', path: '/ac-export/payment/clinic', access: Access.AcExport, form: 'FrmPostingPay2Cln', group: 'Payment', phase: 9 },
      { label: 'Member', path: '/ac-export/payment/member', access: Access.AcExport, form: 'FrmPostingPay2MBM', group: 'Payment', phase: 9 },
      { label: 'Grp Company', path: '/ac-export/bordx/grp-company', access: Access.AcExport, form: 'frmPostingBordxtoGrpComp', group: 'Bordx', desktop: 'hidden', phase: 9 },
      { label: 'Grp Company (GST)', path: '/ac-export/bordx/grp-company-gst', access: Access.AcExport, form: 'frmPostingBordxtoGrpGST', group: 'Bordx', phase: 9 },
      { label: 'Grp Company', path: '/ac-export/receipt/grp-company', access: Access.AcExport, form: 'frmPostingRecGrpCompany', group: 'Receipt', phase: 9 },
      { label: 'SUN System GST', path: '/ac-export/sun/gst', access: Access.AcExport, form: 'frmSUNGSTRpt', group: 'SUN System', phase: 9 },
      { label: 'SUN System Report', path: '/ac-export/sun/report', access: Access.AcExport, form: 'frmSUNSystem', group: 'SUN System', phase: 9 },
      { label: 'SUN System Allocation', path: '/ac-export/sun/allocation', access: Access.AcExport, form: 'frmSUNSystemAllocation, frmSUNSystemSearchAcc', group: 'SUN System', phase: 9 },
      { label: 'SUN System Allocation (New)', path: '/ac-export/sun/allocation-new', access: Access.AcExport, form: 'frmSUNSystemAllocation_New', group: 'SUN System', desktop: 'never-opened', phase: 9 },
    ],
  },
  {
    key: 'tracking',
    label: 'Tracking',
    icon: 'activity',
    path: '/tracking',
    items: [
      { label: 'Clinic', path: '/tracking/clinic', access: Access.ClinicTracking, form: 'FrmTracking', phase: 10 },
      { label: 'Grp. Com. Clinic', path: '/tracking/grp-company-clinic', access: Access.ClinicTracking, form: 'FrmTracking2', phase: 10 },
    ],
  },
  {
    key: 'maintenance',
    label: 'Maintenance',
    icon: 'tool',
    path: '/maintenance',
    items: [
      { label: 'Entity', path: '/maintenance/entity', access: Access.MaintenanceEntity, form: 'frmMaintenanceEntity, frmSearchCLN', phase: 1, migrated: true },
      { label: 'Plan', path: '/maintenance/plan', access: Access.MaintenancePlan, form: 'frmMaintenancePlan', phase: 1, migrated: true },
      { label: 'Case', path: '/maintenance/case', access: Access.MaintenanceCase, form: 'frmMaintenanceOthers', phase: 1, migrated: true },
      { label: 'Calculation', path: '/maintenance/calculation', access: Access.MaintenanceCalculation, form: 'frmMaintenanceCal', phase: 1, migrated: true },
      { label: 'LG Maintenance', path: '/maintenance/lg', access: Access.MaintenanceLg, form: 'frmLGMaintenance', phase: 1, migrated: true },
      { label: 'Member', path: '/maintenance/benefit/member', access: Access.MaintenanceBenefitMember, form: 'frmMaintenanceBenefits', group: 'Benefit', phase: 1, migrated: true },
      { label: 'Clinic', path: '/maintenance/benefit/clinic', access: Access.MaintenanceBenefitClinic, form: 'frmMaintenanceBenefitsCln', group: 'Benefit', phase: 1, migrated: true },
      { label: 'Clinic (new)', path: '/maintenance/benefit/clinic-new', access: Access.MaintenanceBenefitClinic, form: 'frmMaintenanceBenefitsClnNew', group: 'Benefit', desktop: 'never-opened', phase: 1, migrated: true },
      { label: 'Dental', path: '/maintenance/benefit/dental', access: Access.MaintenanceBenefitMember, form: 'frmMaintenanceBNFDental', group: 'Benefit', phase: 1, migrated: true },
      { label: 'Check Item', path: '/maintenance/check-item', access: Access.MaintenanceCheckItem, form: 'frmCheckingItems', group: 'Other', phase: 1, migrated: true },
      { label: 'DataTransfer', path: '/maintenance/data-transfer', access: Access.MaintenanceCheckItem, form: '—', group: 'Other', desktop: 'nothing', phase: 1 },
      { label: 'Test Form', path: '/maintenance/test-form', access: Access.MaintenanceCheckItem, form: 'frmTest', group: 'Other', phase: 11 },
    ],
  },
]

/** Items that need a screen (excludes desktop menu items without a handler and without a form). */
export const buildable = (m: MenuModule) => m.items.filter((i) => i.desktop !== 'nothing')

export const DESKTOP_STATE_TEXT: Record<DesktopState, string> = {
  visible: 'In the desktop menu',
  hidden: 'Hidden in the desktop',
  disabled: 'Disabled in the desktop',
  'never-opened': 'Form in the project, never opened by the desktop',
  nothing: 'Menu item without a handler (nothing to build)',
}
