namespace Cis.Domain.Security;

/// <summary>
/// Menu access positions in USR.USRAccess (legacy VB6 <c>Medix.CheckAccess(n)</c> = character n is "1"). Values come from the
/// <c>*_Click</c> handlers of the CIS <c>MDIMain.frm</c>. Forms the desktop never opens (hidden / no handler / commented out) are gated
/// like their closest sibling; each such case says so. Keep in step with <c>CisUI/src/app/core/auth/access.ts</c>
/// (a test compares the two files).
/// </summary>
public static class AccessRight
{
    // ---- Membership ----
    /// <summary>Membership > MBM Adjustment (FrmMBMCLNAdj), MBM Enquiry (FrmMBMCLNEnq), MBM Topup Adj (FrmMBMTopupLmt).</summary>
    public const int MbmAdjustment = 113;
    /// <summary>Membership > MBM Enquiry (MBMEnquiry_Click).</summary>
    public const int MbmEnquiry = 113;
    /// <summary>Membership > MBM Topup Adj (mnMBMTopup_Click).</summary>
    public const int MbmTopup = 113;
    /// <summary>Membership > MBM Tracking / (Simple) (mnuMBMTracking_Click, MBMTrackingSimple_Click); FrmMBMTracking (never opened) likewise.</summary>
    public const int MbmTracking = 114;
    /// <summary>Membership > MBM Tracking (Jupiter) / (Jupiter - Simple) (mnuMBMTracking2_Click, MBMTrackingJupSimple_Click).</summary>
    public const int MbmTrackingJupiter = 160;

    // ---- MCO Fees ----
    /// <summary>MCOFees > MCO Invoice / GST / GST Zeroised (mnuMCOInvoice_Click, GSTInvoice_Click, mnuMCOFeeGSTZeroised_Click).</summary>
    public const int McoInvoice = 115;
    /// <summary>MCOFees > Fee Listing / Refund Listing (mnuFeeListing_Click, mnuFeeRefund_Click).</summary>
    public const int McoFeeListing = 116;

    // ---- Invoice ----
    /// <summary>Invoice > Clinic > Clinic (mnuInvProcess_Click) and the three Email Notification items (mnuEmail*_Click).</summary>
    public const int InvoiceClinic = 117;
    /// <summary>Invoice > Clinic > Member (mnuInvMember_Click).</summary>
    public const int InvoiceMember = 118;
    /// <summary>
    /// Invoice > Adjustment, Adjustment (After Posting), Enquiry, Batch Transfer, Invoice Update, OP Batch/Invoice, the Claim ASCII items,
    /// LG > Generate LG and Audit review / Audit Tracking (all check 119).
    /// </summary>
    public const int InvoiceAdjustment = 119;
    /// <summary>Invoice > Clinic > Adjustment (After Bordx) (mnuINVAdjAfterBordx_Click).</summary>
    public const int InvoiceAdjustmentAfterBordx = 179;
    /// <summary>Invoice > Email Notification (Payment / Providers / Rejection).</summary>
    public const int EmailNotification = 117;
    /// <summary>Invoice > General Exclusions (mnuGeneralExclusions_Click).</summary>
    public const int GeneralExclusions = 133;
    /// <summary>Invoice > General Exclusions (Enquiry) (mnuGeneralExclusionsEnq_Click).</summary>
    public const int GeneralExclusionsEnquiry = 113;
    /// <summary>Invoice > Claim ASCII (Bintulu, Toshiba, Singer, Utusan, OMNI, TIMS, CIMB) and the OMNI providers form.</summary>
    public const int ClaimAscii = 119;
    /// <summary>Invoice > Case Transfer (mnuCaseTransfer_Click).</summary>
    public const int CaseTransfer = 162;
    /// <summary>Invoice > Tracking (mnuINVTracking_Click).</summary>
    public const int InvoiceTracking = 161;
    /// <summary>Invoice > Tracking (Jupiter) (mnuINVTracking2_Click).</summary>
    public const int InvoiceTrackingJupiter = 162;
    /// <summary>Invoice > MMA/CPT (mnuMMACPT_Click).</summary>
    public const int MmaCpt = 150;
    /// <summary>Worksheet Registration (frmNewWSReg, never opened): gated like Invoice > Clinic, which opens the worksheet pop-up.</summary>
    public const int WorksheetRegistration = 117;
    /// <summary>Worksheet Adjustment (frmNewWSAjd, never opened): gated like Invoice > Adjustment.</summary>
    public const int WorksheetAdjustment = 119;

    // ---- Claim ----
    /// <summary>Claim > Pass to Claim (mnuPass2Claim_Click).</summary>
    public const int PassToClaim = 163;

    // ---- Claim Bordx ----
    /// <summary>Claim Bordx > Worksheet > Insert Worksheet / Insert Invoice and ClaimBordxDate (122).</summary>
    public const int BordxInsert = 122;
    /// <summary>Claim Bordx > Worksheet > Remove Worksheet (mnuBordxRevise_Click).</summary>
    public const int BordxRemove = 123;
    /// <summary>Claim Bordx > Update Claims Data and every Print Bordx item (124), incl. BSN-PrintSummary (handler commented out).</summary>
    public const int BordxPrint = 124;
    /// <summary>Claim Bordx > Bordx Delivery > Send to Grp. Company (mnuToPayor_Click); Send to MNRB (never opened) likewise.</summary>
    public const int BordxSendToPayor = 125;
    /// <summary>Claim Bordx > Bordx Delivery > Returned Fr Grp. Company (mnuFrPayor_Click).</summary>
    public const int BordxReturnedFromPayor = 126;
    /// <summary>Claim Bordx > Bordx Delivery > Send to A/C (mnuSentAC_Click).</summary>
    public const int BordxSendToAc = 127;
    /// <summary>Claim Bordx > Bor Acc Period (mnuBorAccPeriod_Click).</summary>
    public const int BordxAcPeriod = 128;
    /// <summary>Claim Bordx > Monitoring > Bordx Control (mnuBordxCtrl_Click).</summary>
    public const int BordxControl = 129;
    /// <summary>Claim Bordx > Monitoring > Outstanding Bordx (mnuOSBordx_Click).</summary>
    public const int BordxOutstanding = 130;
    /// <summary>Claim Bordx > Monitoring > Bordx &amp; Worksheet (mnuBordxTracking_Click).</summary>
    public const int BordxWorksheet = 131;
    /// <summary>Claim Bordx > Claim Tracking / (Simple) / (SUN System) (132).</summary>
    public const int ClaimTracking = 132;
    /// <summary>Claim Bordx > Claim Tracking 2 (mnuClaimTrack2_Click checks 120).</summary>
    public const int ClaimTracking2 = 120;
    /// <summary>Claim Bordx > Claim Tracking (Jupiter) / (Jupiter - Simple) (163).</summary>
    public const int ClaimTrackingJupiter = 163;
    /// <summary>Claim Bordx > Claim Tracking 2 (Jupiter) (mnuClaimTrack4_Click).</summary>
    public const int ClaimTracking2Jupiter = 164;

    // ---- LG / Audit ----
    /// <summary>LG > Generate LG (mnuGenLG_Click).</summary>
    public const int GenerateLg = 119;
    /// <summary>Audit review > Audit review / Audit Tracking (mnuAudit_Click, mnuAudittracking_Click).</summary>
    public const int AuditReview = 119;

    // ---- Payment ----
    /// <summary>Payment > To Clinic (mnuPaymentCLN_Click).</summary>
    public const int PaymentClinic = 133;
    /// <summary>Payment > To Clinic - Invoice and FFS Payment (121).</summary>
    public const int PaymentClinicInvoice = 121;
    /// <summary>Payment > To Member (mnuPaymentMBM_Click); frmPaymentMember (never opened) likewise.</summary>
    public const int PaymentMember = 134;
    /// <summary>Payment > Tracking (mnuPaymentTracking_Click).</summary>
    public const int PaymentTracking = 165;
    /// <summary>Payment > Tracking (Jupiter) (mnuPaymentTracking2_Click).</summary>
    public const int PaymentTrackingJupiter = 167;
    /// <summary>Payment > FFS Payment (mnuPaymentFFS_Click).</summary>
    public const int PaymentFfs = 121;

    // ---- Receipt ----
    /// <summary>Receipt > Fr Member (mnuReceiptMBM_Click).</summary>
    public const int ReceiptMember = 135;
    /// <summary>Receipt > Fr Grp Company, Upload Receipt, Upload Receipt (Multiple) (136); Receipt Fr MNRB (never opened) likewise.</summary>
    public const int ReceiptPayor = 136;
    /// <summary>Receipt > Fr FFS (mnuReceiptFFS_Click).</summary>
    public const int ReceiptFfs = 121;

    // ---- A/C Export ----
    /// <summary>Every A/C Export item (invoice / payment / bordx / receipt posting, SUN System GST / Report / Allocation) checks 120.</summary>
    public const int AcExport = 120;

    // ---- Tracking ----
    /// <summary>Tracking > Clinic / Grp. Com. Clinic (mnuClinic_Click, mnuGrpComCLN_Click).</summary>
    public const int ClinicTracking = 144;

    // ---- Maintenance ----
    /// <summary>Maintenance > Entity (mnuEntity_Click).</summary>
    public const int MaintenanceEntity = 137;
    /// <summary>Maintenance > Plan and LGMaintenance (mnuPlan_Click, frmLgMaintainace_Click).</summary>
    public const int MaintenancePlan = 138;
    /// <summary>Maintenance > Case (MaintenanceOthers_Click).</summary>
    public const int MaintenanceCase = 139;
    /// <summary>Maintenance > Calculation (mnuCalculation_Click).</summary>
    public const int MaintenanceCalculation = 140;
    /// <summary>Maintenance > LGMaintenance (frmLgMaintainace_Click checks 138).</summary>
    public const int MaintenanceLg = 138;
    /// <summary>Maintenance > Benefit > Member / Dental (mnuMember_Click, mnuBnfDental_Click).</summary>
    public const int MaintenanceBenefitMember = 141;
    /// <summary>Maintenance > Benefit > Clinic (mnuBnfClinic_Click); frmMaintenanceBenefitsClnNew (never opened) likewise.</summary>
    public const int MaintenanceBenefitClinic = 142;
    /// <summary>Maintenance > Check Item (mnuCheckItem_Click).</summary>
    public const int MaintenanceCheckItem = 143;

    /// <summary>True when character <paramref name="position"/> (1-based) of the access string is "1".</summary>
    public static bool Has(string? access, int position) =>
        access is not null && position >= 1 && access.Length >= position && access[position - 1] == '1';
}
