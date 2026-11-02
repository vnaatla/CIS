using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Invoice;

/// <summary>How much of a loaded case may be changed (CmdEdit_Click).</summary>
public enum AdjustmentEditMode
{
    /// <summary>Everything.</summary>
    Full,
    /// <summary>Posted to SUN: only the diagnoses and the remark.</summary>
    DiagnosisOnly,
    /// <summary>Bordered or paid: no change.</summary>
    None,
}

/// <summary>The adjustment's own fields (status, pending, repudiation, BTBG, GST, medication charges, notification).</summary>
public sealed record AdjustmentExtra
{
    /// <summary>Case status O / C / D.</summary>
    public string? Status { get; init; }
    /// <summary>Claimability A / R / C / S.</summary>
    public string? Claimability { get; init; }
    /// <summary>Pending Y / N.</summary>
    public string? Pending { get; init; }
    /// <summary>Pay to C / M.</summary>
    public string? PayTo { get; init; }
    /// <summary>Accepted (bordereaux memo) Y / N.</summary>
    public string? MemoClaim { get; init; }
    /// <summary>Referral letter Y / N.</summary>
    public string? RefLetter { get; init; }
    /// <summary>Pharmacy Y / N.</summary>
    public string? Pharmacy { get; init; }
    /// <summary>NP / NPS clinic name.</summary>
    public string? NpClinic { get; init; }
    /// <summary>Repudiation reason and amount.</summary>
    public string? RepReason { get; init; }
    /// <summary>Repudiation amount.</summary>
    public string? RepAmount { get; init; }
    /// <summary>Management notes (CASMgmtNotes).</summary>
    public string? MgmtNotes { get; init; }
    /// <summary>BTBG amount and remark.</summary>
    public decimal? BtbgAmount { get; init; }
    /// <summary>BTBG remark.</summary>
    public string? BtbgRemark { get; init; }
    /// <summary>frmCasePending: checked reason codes, questions to the hospital, additional remark.</summary>
    public IReadOnlyList<string> PendingReasonIds { get; init; } = [];
    /// <summary>Questions (QA to hospital).</summary>
    public IReadOnlyList<string> Questions { get; init; } = [];
    /// <summary>Additional pending remark.</summary>
    public string? AddPendRemark { get; init; }
    /// <summary>Names of the checked reasons (the notification's pending reasons).</summary>
    public IReadOnlyList<string> PendingReasonNames { get; init; } = [];
    /// <summary>GST of A1..A33 (index) and B5..B8.</summary>
    public IReadOnlyList<decimal?> AGst { get; init; } = [];
    /// <summary>GST of B1..B8 (index; B5–B8 used).</summary>
    public IReadOnlyList<decimal?> BGst { get; init; } = [];
    /// <summary>Medication charges 1–7.</summary>
    public IReadOnlyList<decimal?> MedCharges { get; init; } = [];
    /// <summary>Rejection / pending e-mail: the additional notes (null = the user cancelled) and an e-mail when the member has none.</summary>
    public string? NotifyNotes { get; init; }
    /// <summary>E-mail typed for a member without one.</summary>
    public string? NotifyEmail { get; init; }
}

/// <summary>Update request.</summary>
/// <param name="CaseNo">Case.</param>
/// <param name="Entry">Items, diagnoses, visit (shared with the invoice screens).</param>
/// <param name="Extra">Adjustment fields.</param>
public record AdjustmentRequest(string CaseNo, ClinicInvoiceEntry Entry, AdjustmentExtra Extra);

/// <summary>Lengths.</summary>
public sealed class AdjustmentRequestValidator : Common.RequestValidator<AdjustmentRequest>
{
    /// <summary>Defines the rules.</summary>
    public AdjustmentRequestValidator()
    {
        RuleFor(r => r.CaseNo).NotEmpty().MaximumLength(20);
        RuleFor(r => r.Entry).NotNull();
        RuleFor(r => r.Extra).NotNull();
        RuleFor(r => r.Entry.InvoiceNo).MaximumLength(50);
        RuleFor(r => r.Entry.Remark).MaximumLength(500);
        RuleFor(r => r.Extra.RepReason).MaximumLength(1000);
        RuleFor(r => r.Extra.MgmtNotes).MaximumLength(2000);
        RuleFor(r => r.Extra.BtbgRemark).MaximumLength(500);
        RuleFor(r => r.Extra.AddPendRemark).MaximumLength(2000);
        RuleFor(r => r.Entry.A.Count).LessThanOrEqualTo(34);
        RuleFor(r => r.Entry.B.Count).LessThanOrEqualTo(9);
    }
}

/// <summary>Audit review request.</summary>
/// <param name="Query">Query.</param>
/// <param name="Answer">Answer.</param>
/// <param name="CostAvoidance">Discount received.</param>
/// <param name="Remarks">Remarks.</param>
/// <param name="AuditBy">Audit by.</param>
/// <param name="AuditDate">dd/mm/yyyy.</param>
public record AuditReviewRequest(string? Query, string? Answer, string? CostAvoidance, string? Remarks, string? AuditBy, string? AuditDate);

/// <summary>Lengths.</summary>
public sealed class AuditReviewRequestValidator : Common.RequestValidator<AuditReviewRequest>
{
    /// <summary>Defines the rules.</summary>
    public AuditReviewRequestValidator()
    {
        RuleFor(r => r.Query).MaximumLength(4000);
        RuleFor(r => r.Answer).MaximumLength(4000);
        RuleFor(r => r.Remarks).MaximumLength(4000);
        RuleFor(r => r.AuditBy).MaximumLength(100);
        RuleFor(r => r.CostAvoidance).MaximumLength(50);
    }
}

/// <summary>A loaded case.</summary>
public sealed record AdjustmentCase
{
    /// <summary>SearchCASOP row.</summary>
    public Dictionary<string, object?> Case { get; init; } = [];
    /// <summary>Claim row.</summary>
    public Dictionary<string, object?>? Claim { get; init; }
    /// <summary>Payments exist / posted to SUN / bordered.</summary>
    public bool Paid { get; init; }
    /// <summary>Posted to SUN.</summary>
    public bool Posted { get; init; }
    /// <summary>INVBordxed.</summary>
    public bool Bordxed { get; init; }
    /// <summary>What Edit allows and its message.</summary>
    public AdjustmentEditMode EditMode { get; init; }
    /// <summary>Edit message.</summary>
    public string? EditMessage { get; init; }
    /// <summary>The latest pending reason codes.</summary>
    public List<string> PendingReasonIds { get; init; } = [];
    /// <summary>The reasons of the case's pay-to category (empty when the case already has reasons, as the desktop).</summary>
    public List<LookupOption> PendingReasons { get; init; } = [];
    /// <summary>SP_SpecialistCopayamt (replaces the co-pay when present).</summary>
    public decimal? SpecialistCoPay { get; init; }
    /// <summary>The member's e-mail.</summary>
    public string Email { get; init; } = "";
    /// <summary>The plan's special remark (shown on Edit).</summary>
    public string SpecialRemark { get; init; } = "";
    /// <summary>PLN.PLNBTBG.</summary>
    public bool Btbg { get; init; }
    /// <summary>The member panel.</summary>
    public InvoiceMember? Member { get; init; }
    /// <summary>The stored worksheet.</summary>
    public StoredWorksheet Worksheet { get; init; } = new([], null);
    /// <summary>Messages after Go (same-visit cases, already audited).</summary>
    public List<string> Messages { get; init; } = [];
}

/// <summary>Update result.</summary>
/// <param name="Message">The desktop's final message.</param>
/// <param name="OpBatchNo">OP batch after the update.</param>
/// <param name="Warnings">Non-blocking messages shown before the save.</param>
/// <param name="NewCaseNo">After Posting: the created sub-case.</param>
public record AdjustmentResult(string Message, string OpBatchNo, List<string> Warnings, string? NewCaseNo = null);

/// <summary>Values the adjustment save writes.</summary>
/// <param name="Request">Request.</param>
/// <param name="Loaded">The case before the change.</param>
/// <param name="Amounts">A1..A33.</param>
/// <param name="Total">Total bill.</param>
/// <param name="ActCoPay">lblActCoPay.</param>
/// <param name="Worksheet">The new worksheet.</param>
/// <param name="RepAction">CASRepAction (pending reason history).</param>
/// <param name="UserCode">User.</param>
/// <param name="Today">Today.</param>
/// <param name="McoProcFee">MCO processing fee.</param>
/// <param name="ClinicBill">lblCLNBillAmt.</param>
/// <param name="AfterBordx">frmInvoiceAdjAftBordx: only the case (without QA / pending note) and the invoice are saved.</param>
/// <param name="NewCase">frmInvoiceAdjAfterPost: the sub-case created from the loaded (master) case.</param>
public record AdjustmentSave(AdjustmentRequest Request, AdjustmentCase Loaded, decimal[] Amounts, decimal Total, decimal ActCoPay, WorksheetResult Worksheet,
    string RepAction, string UserCode, DateTime Today, decimal McoProcFee, decimal ClinicBill, bool AfterBordx = false, string? NewCase = null);

/// <summary>Data access of the adjustment (<c>cis.InvAdj_*</c> and the desktop's INVAdj procedures).</summary>
public interface IInvoiceAdjustmentRepository
{
    /// <summary>cis.InvAdj_Load: case, claim, paid / posted, reason ids, specialist co-pay, e-mail, plan remark.</summary>
    Task<(Dictionary<string, object?>? Case, Dictionary<string, object?>? Claim, bool Paid, bool Posted, List<string> ReasonIds, decimal? CoPay, string Email,
        string SpecialRemark, bool Btbg)> LoadAsync(IDbSession session, string casNumber);

    /// <summary>Pending reasons of a category.</summary>
    Task<List<LookupOption>> PendingReasonsAsync(IDbSession session, string category);

    /// <summary>spMemberduplicatecases_Identify.</summary>
    Task<List<string>> SameVisitCasesAsync(IDbSession session, string casNumber);

    /// <summary>spMemberduplicateInvoices_Identify: the duplicate cases (empty = none).</summary>
    Task<List<string>> DuplicateInvoicesAsync(IDbSession session, string casNumber, DateTime invoiceDate, string billAmt, string invoiceNo, string patientName,
        string patientIc, string userCode, string coverId);

    /// <summary>sp_chkDoubleCaseEntry's isExist.</summary>
    Task<string> DoubleEntryAsync(IDbSession session, string mbmNumber, string coverId, string clinicCode, string invoiceNo, string visitTime, DateTime visit,
        decimal totalActual, string casNumber);

    /// <summary>The member's other claims (CheckExceedLimit).</summary>
    Task<List<InvClaim>> OtherClaimsAsync(IDbSession session, string mbmNumber, int mode, string covId, string casNumber);

    /// <summary>The worksheet tables of the case (for the contracted rates).</summary>
    Task<List<Dictionary<string, object?>?>> WorksheetAsync(IDbSession session, string casNumber);

    /// <summary>DiagnosisUpdate.</summary>
    Task DiagnosisAsync(IDbSession session, string casNumber, IReadOnlyList<string?> diagnoses, string other, string remark);

    /// <summary>SaveRecord (in the open transaction).</summary>
    Task SaveAsync(IDbSession session, AdjustmentSave save);

    /// <summary>UpdateRecord / UpdateEmptyOPbatchno.</summary>
    Task BatchAsync(IDbSession session, string casNumber, string? opBatchNo, string userCode, DateTime today);

    /// <summary>The next OC batch.</summary>
    Task<string> NextBatchAsync(IDbSession session);

    /// <summary>sp_EmailNotifyPendingReject_LatestModified: true when sent.</summary>
    Task<bool> NotifyAsync(IDbSession session, string casNumber, string notes, string pendingReasons, string email);

    /// <summary>Audit review.</summary>
    Task AuditAsync(IDbSession session, string casNumber, AuditReviewRequest request, DateTime auditDate);

    /// <summary>CASClaimsApp = Y.</summary>
    Task ClaimsAppAsync(IDbSession session, string casNumber);

    /// <summary>Void bordereaux: the KIV number.</summary>
    Task<string> VoidBordxAsync(IDbSession session, string casNumber, string groupCompany, string userCode, DateTime today);
}

/// <summary>Invoice > Clinic > Adjustment (frmInvoiceAdj).</summary>
public interface IInvoiceAdjustmentService
{
    /// <summary>CmdGo_Click.</summary>
    Task<AdjustmentCase> LoadAsync(string? casNumber);

    /// <summary>The worksheet of the changed case (Compute).</summary>
    Task<WorksheetResult> ComputeAsync(AdjustmentRequest request, string userCode);

    /// <summary>CmdUpdate_Click.</summary>
    Task<AdjustmentResult> UpdateAsync(AdjustmentRequest request, string userCode);

    /// <summary>cmdauditreviewupdate_Click.</summary>
    Task<string> AuditAsync(string? casNumber, AuditReviewRequest request);

    /// <summary>cmdCLMBordxApp_Click.</summary>
    Task ClaimsAppAsync(string? casNumber);

    /// <summary>cmdvoidbordx_Click.</summary>
    Task<string> VoidBordxAsync(string? casNumber, string userCode);

    /// <summary>frmInvoiceAdjAftBordx CmdGo_Click: a bordered case is editable unless posted to SUN.</summary>
    Task<AdjustmentCase> LoadAfterBordxAsync(string? casNumber);

    /// <summary>frmInvoiceAdjAftBordx CmdUpdate_Click: the case and its invoice only.</summary>
    Task<AdjustmentResult> UpdateAfterBordxAsync(AdjustmentRequest request, string userCode);

    /// <summary>frmInvoiceAdjAfterPost CmdGo_Click: only a case posted to SUN can get a sub-case.</summary>
    Task<AdjustmentCase> LoadAfterPostAsync(string? casNumber);

    /// <summary>frmInvoiceAdjAfterPost CmdCreate_Click: the new sub-case number.</summary>
    Task<AdjustmentResult> CreateAfterPostAsync(AdjustmentRequest request, string userCode);
}

/// <inheritdoc />
public sealed class InvoiceAdjustmentService(IDbSessions sessions, IInvoiceAdjustmentRepository repository, IClinicInvoiceRepository clinics,
    IClinicInvoiceService invoices, IWorksheetRepository worksheet, TimeProvider clock) : IInvoiceAdjustmentService
{
    /// <summary>Desktop texts.</summary>
    public const string NotFound = "Not Record Found ! ", PostedOnly = "This case posted into SUN System Only can amend Diagnosis and Remark ! ",
        Bordered = "This case already Bordx ! ", PaidDone = "This case already Payment done ! ", Audited = "This case already audited..!",
        EmailMissing = "Member Email Address is Missing. Rejection/Pending Email Cannot Be sent.Please check with Memebership Team! ",
        DoubleEntry = "This Case Number information seems created before! Please check this Member Case History to avoid double entry!";

    /// <inheritdoc />
    public async Task<AdjustmentCase> LoadAsync(string? casNumber)
    {
        var cas = Case(casNumber);
        AdjustmentCase loaded;
        await using (var s = await sessions.OpenAsync()) loaded = await Load(s, cas);
        var messages = loaded.Messages;
        try
        {
            var member = await invoices.MemberAsync(T(loaded.Case, "MBMnumber"), T(loaded.Case, "CoverID"));
            messages.AddRange(member.Warnings);
            return loaded with { Member = member };
        }
        catch (BusinessRuleException e)
        {
            messages.Add(e.Message);
            return loaded;
        }
    }

    /// <inheritdoc />
    public async Task<WorksheetResult> ComputeAsync(AdjustmentRequest r, string userCode)
    {
        await using var s = await sessions.OpenAsync();
        var loaded = await Load(s, Case(r.CaseNo));
        var member = await invoices.MemberAsync(T(loaded.Case, "MBMnumber"), T(loaded.Case, "CoverID"));
        return (await Line(s, r, loaded, member, userCode)).Worksheet;
    }

    /// <inheritdoc />
    public async Task<AdjustmentResult> UpdateAsync(AdjustmentRequest r, string userCode)
    {
        var cas = Case(r.CaseNo);
        var user = userCode.Trim().ToUpperInvariant();
        var today = clock.GetLocalNow().Date;
        var e = r.Entry;
        var x = r.Extra;
        var warnings = new List<string>();
        await using var s = await sessions.BeginAsync();
        var loaded = await Load(s, cas);
        switch (loaded.EditMode)
        {
            case AdjustmentEditMode.None:
                throw new BusinessRuleException(loaded.EditMessage!);
            case AdjustmentEditMode.DiagnosisOnly:
                await repository.DiagnosisAsync(s, cas, e.Diagnoses, (e.DiagnosisOther ?? "").Trim(), (e.Remark ?? "").Trim());
                await s.CommitAsync();
                return new AdjustmentResult("", T(loaded.Case, "OPBatchNo"), warnings);
        }
        var member = await invoices.MemberAsync(T(loaded.Case, "MBMnumber"), T(loaded.Case, "CoverID"));
        var m = member.Member;
        var line = await Line(s, r, loaded, member, userCode);
        var number = m.MbmNumber;
        var ubbs = number.StartsWith("UB", StringComparison.Ordinal) || number.StartsWith("BS", StringComparison.Ordinal);
        string status = F(x.Status), claim = F(x.Claimability), pending = F(x.Pending), payTo = F(x.PayTo), memo = F(x.MemoClaim);
        bool rejection = status == "D" && claim == "C" && pending == "N", pendingOpen = status == "O" && claim == "A" && pending == "Y";
        if (ubbs && payTo == "M" && (rejection || pendingOpen) && loaded.Email.Length == 0) warnings.Add(EmailMissing);

        var clinic = (e.ClinicCode ?? "").Trim().ToUpperInvariant();
        var clinicName = clinic is "NP" or "NPS" ? "x" : (await clinics.ClinicAsync(s, clinic))?.Name ?? "";
        var otherDiag = e.Diagnoses.Count > 4 && !string.IsNullOrWhiteSpace(e.Diagnoses[4])
            && (await clinics.DiagnosisAsync(s, e.Diagnoses[4]!.Trim()))?.OtherIndicator?.Trim() == "Y";
        var bill = ClinicInvoiceRules.NumOrNull(e.ClinicBillAmount);
        var hasReasons = x.PendingReasonIds.Count > 0 || x.Questions.Count > 0 || !string.IsNullOrWhiteSpace(x.AddPendRemark);
        string? message = null;
        if (string.IsNullOrWhiteSpace(e.MbmNumber)) message = "Please keyin Member No ! ";
        else if (string.IsNullOrWhiteSpace(e.VisitDate)) message = "Please Enter Visit Date! ";
        else if (clinicName.Length == 0) message = "Please keyin the Clinic code ! ";
        else if (clinic == "NP" && string.IsNullOrWhiteSpace(x.NpClinic)) message = "Please keyin the NP Clinic ! ";
        else if (clinic == "NPS" && string.IsNullOrWhiteSpace(x.NpClinic)) message = "Please keyin the NPS Clinic ! ";
        else if (otherDiag && string.IsNullOrWhiteSpace(e.DiagnosisOther)) message = "Please Keyin Others Diagnosis ! ";
        else if (payTo == "C" && clinic == "NP") message = "Invalid Clinic Code";
        else if (bill is null && memo == "Y") message = "Please Enter Bill Amount ! ";
        else if ((bill ?? 0) != line.Total) message = "Please check Bill Amount! ";
        else if (clinic == "NPS" && string.IsNullOrWhiteSpace(x.RefLetter)) message = "Please enter status for referral letter";
        else if (!string.IsNullOrWhiteSpace(e.InvoiceNo) && ClinicInvoiceRules.Date(e.InvoiceDate) is null or { Year: 1900, Month: 1, Day: 1 })
            message = "Please enter invoice date";
        else if (status == "O" && claim.Length > 0 && T(loaded.Case, "GLIssued") != "Y"
                 && await ExceedsLimit(s, cas, member, line, bill ?? 0, clinicType: T(loaded.Case, "CLNType") is { Length: > 0 } ct ? ct : "G"))
            message = MemberInvoiceService.OverBalance;
        else if (pending == "Y" && status != "O") message = "Please check case status.. ! ";
        else if (line.Total == 0 && status == "O" && pending != "Y") message = "Total Invoice Amount cannot be Zero ! ";
        else if (status == "O" && claim == "A" && pending == "Y" && !hasReasons) message = "Case is in Pending State, Please select Pending Reasons...! ";
        else if (pending == "N" && hasReasons) message = "Case pending status changed to 'N', but pending details not cleared! Please check ";
        else if (number.StartsWith("UB", StringComparison.Ordinal) && payTo == "M" && ClinicInvoiceRules.Date(e.InvoiceDate) is { } invDate
                 && !string.IsNullOrWhiteSpace(e.InvoiceNo) && !string.IsNullOrWhiteSpace(e.ClinicBillAmount)
                 && await repository.DuplicateInvoicesAsync(s, cas, invDate, e.ClinicBillAmount!.Trim(), e.InvoiceNo!.Trim(), m.PatientName, m.PatientIc, user, m.CoverId)
                     is [_, ..] dups)
            message = "Duplicate case found: " + string.Join(",", dups);
        else if (ubbs && rejection && payTo == "M" && string.IsNullOrWhiteSpace(x.RepReason)) message = "Repudation Reason cannot be Empty.Please fill Repudation Reason..! ";
        else if (ubbs && pendingOpen && payTo == "M" && x.PendingReasonNames.Count == 0) message = "Action cannot be Empty.Please fill Proper Action..! ";
        else if (loaded.Btbg && m.PlanCode is "G018" or "G019" && line.Total != 0 && line.Total - 60 != (x.BtbgAmount ?? 0) && string.IsNullOrWhiteSpace(x.BtbgRemark))
            message = "Please enter BTBG Remark ! ";
        else if (T(loaded.Case, "CLNType") == "S" && T(loaded.Case, "INVPayTo") == "C" && ClinicInvoiceRules.Date(e.VisitDate) is { } visit
                 && await repository.DoubleEntryAsync(s, m.MbmNumber, m.CoverId, clinic, (e.InvoiceNo ?? "").Trim(), (e.VisitTime ?? "").Trim(), visit, bill ?? 0, cas) == "Y")
            message = DoubleEntry;
        if (message is not null) throw new BusinessRuleException(message);

        await repository.SaveAsync(s, line with { Request = r, UserCode = user, Today = today });
        var batch = T(loaded.Case, "OPBatchNo");
        if (!string.IsNullOrWhiteSpace(e.InvoiceNo) && batch.Length == 0 && status == "O" && claim == "A" && pending == "N")
        {
            batch = await repository.NextBatchAsync(s);
            await repository.BatchAsync(s, cas, batch, user, today);
        }
        else if (string.IsNullOrWhiteSpace(e.InvoiceNo))
        {
            await repository.BatchAsync(s, cas, null, user, today);
            batch = "";
        }
        await s.CommitAsync();

        var final = await Notify(r, loaded, ubbs, rejection, pendingOpen, batch);
        if (member.AsoInd == "Y")
        {
            await using var s2 = await sessions.OpenAsync();
            await clinics.AsoUtilisationAsync(s2, cas);
        }
        return new AdjustmentResult(final, batch, warnings);
    }

    /// <summary>The rejection / pending e-mail and the desktop's final message.</summary>
    private async Task<string> Notify(AdjustmentRequest r, AdjustmentCase loaded, bool ubbs, bool rejection, bool pendingOpen, string batch)
    {
        var x = r.Extra;
        if (!(ubbs && F(x.PayTo) == "M")) return "Record updated ! ";
        var reasonChanged = rejection && (x.RepReason ?? "").Trim() != T(loaded.Case, "RepudiationReason");
        var pendingChanged = pendingOpen && string.Join(",", x.PendingReasonIds) != string.Join(",", loaded.PendingReasonIds);
        var sent = false;
        if (reasonChanged || pendingChanged)
        {
            if (x.NotifyNotes is null) return "Record updated ! Since no change, no email is Sent. ";
            var email = loaded.Email.Length > 0 ? "" : (x.NotifyEmail ?? "").Trim();
            if (batch.Length == 0 && (loaded.Email.Length > 0 || email.Length > 0))
            {
                await using var s = await sessions.OpenAsync();
                sent = await repository.NotifyAsync(s, Case(r.CaseNo), x.NotifyNotes.Trim(), pendingChanged ? string.Join(",", x.PendingReasonNames) : "", email);
            }
            else if (loaded.Email.Length == 0) return "Record updated ! Since no change, no email is Sent. ";
        }
        if (sent) return "Record updated and Email Sent! ";
        if ((rejection && reasonChanged) || (pendingOpen && pendingChanged)) return "Record updated..!Email id is empty. No email is Sent. ";
        if (rejection || pendingOpen) return "Record updated..!Since no change No email is Sent. ";
        return "Record updated..! ";
    }

    /// <summary>UpdateLstView + the worksheet + the save values.</summary>
    private async Task<AdjustmentSave> Line(IDbSession s, AdjustmentRequest r, AdjustmentCase loaded, InvoiceMember member, string userCode)
    {
        var e = r.Entry;
        var x = r.Extra;
        var m = member.Member;
        var a = ClinicInvoiceRules.Amounts(e);
        if (member.PlanGroupCompany == "C014") a[29] = Enumerable.Range(1, 4).Any(i => i < e.B.Count && e.B[i] is not null) ? 5 : 0;
        var total = ClinicInvoiceRules.Total(a);
        var actCoPay = loaded.SpecialistCoPay ?? N(loaded.Claim, "CLMCoPayCLNColl");
        if (member.PlanGroupCompany == "C014") actCoPay = a[29];
        var payTo = F(x.PayTo) is { Length: > 0 } p ? p : T(loaded.Case, "INVPayTo");
        var tables = await repository.WorksheetAsync(s, Case(r.CaseNo));
        var stored = InvoiceEnquiryService.Stored(Enumerable.Range(0, WorksheetService.Rows).Select(i => ("", "")).ToList(), tables);
        var clinic = (e.ClinicCode ?? "").Trim().ToUpperInvariant();
        var input = new WorksheetInput
        {
            Caller = WorksheetCaller.Adjustment, PayTo = payTo, ClinicCode = clinic, PlanCode = m.PlanCode, InsCode = m.InsType, HltCode = m.HltCode,
            Items = ClinicInvoiceRules.GridItems(a), NonGpAmount = a[31], Row70BalanceLimit = member.BalanceLimit, BalanceLimit = member.BalanceLimit,
            VisitLimit = member.VisitLimit, AnnualIndicator = m.AnnualIndicator != "N", VisitIndicator = m.VisitIndicator != "N", CoPayCollected = actCoPay,
            CoPay = 0, AdjustedAnnualLimit = e.AdjustedAnnualLimit ?? 0, GpStatus = Enumerable.Range(1, 4).Sum(i => i < e.B.Count ? e.B[i] ?? 0 : 0) > 0,
            McoProcStatus = F(e.Ffs) == "Y", ApprovedOverride = payTo == "M" ? m.BordxFees : null,
            StoredRates = stored.Rows.Select(w => w.ContractedRate).ToList(),
        };
        var ws = WorksheetEngine.Calculate(input, await worksheet.LookupsAsync(s, input), clock.GetLocalNow().Date, userCode);
        var clinicBill = total - a[29];
        var mco = 0m;
        if (payTo == "C" && T(loaded.Case, "CLNFFSStatus") == "Y")
            mco = Case(r.CaseNo).StartsWith("SJ", StringComparison.Ordinal) ? N(loaded.Claim, "CLMMCOProcFee") : N(loaded.Case, "CLNMCOProcfeePerc") * clinicBill / 100;
        return new AdjustmentSave(r, loaded with { Member = member }, a, total, actCoPay, ws, RepAction(x, loaded), userCode, clock.GetLocalNow().Date, mco, clinicBill);
    }

    /// <summary>SaveCase's CASRepAction: a new "|date time-ids" entry when the pending reasons changed.</summary>
    public static string RepAction(AdjustmentExtra x, AdjustmentCase loaded) => RepAction(x, loaded, DateTime.Now);

    /// <summary>SaveCase's CASRepAction at <paramref name="now"/>.</summary>
    public static string RepAction(AdjustmentExtra x, AdjustmentCase loaded, DateTime now)
    {
        var ids = F(x.Pending) == "N" ? "" : string.Join(",", x.PendingReasonIds);
        var existing = string.Join(",", loaded.PendingReasonIds);
        var stamp = "|" + now.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture) + "-" + ids;
        var numericExisting = existing.Replace(",", "", StringComparison.Ordinal) is { Length: > 0 } digits && digits.All(char.IsDigit);
        if (F(x.Pending) == "N" && x.PendingReasonIds.Count > 0) return stamp;
        if (numericExisting) return ids == existing ? "" : stamp;
        if (F(x.Pending) == "N") return "";
        return ids == existing ? "" : stamp;
    }

    /// <summary>CheckExceedLimit.</summary>
    private async Task<bool> ExceedsLimit(IDbSession s, string cas, InvoiceMember member, AdjustmentSave line, decimal bill, string clinicType)
    {
        if (ClinicInvoiceRules.NumOrNull(member.AnnualLimit) is not { } annual) return false;
        var a = line.Amounts;
        if (member.AsoInd == "Y")
        {
            var finalApprove = line.Worksheet.RowValues.TryGetValue(161, out var f) && f is decimal d ? d : 0;
            return ClinicInvoiceRules.Num(member.BalanceLimit) + finalApprove - (bill - (a[32] + a[33])) <= 0;
        }
        var m = member.Member;
        bool maternity = member.MaternityInd == "Y", dental = member.DentalInd == "Y";
        var mode = ClinicInvoiceRules.ClaimMode(m.SuppLimitType == "E" ? "C" : m.SuppLimitType, m.CoverId);
        var claims = mode is { } md ? await repository.OtherClaimsAsync(s, m.MbmNumber, md, m.CoverId, cas) : [];
        var specBranch = member.SpecLimit != "N/A";
        var (other, spec) = ClinicInvoiceRules.Utilisation(claims, maternity, dental, specBranch);
        var net = bill - ((dental ? a[7] : 0) + (maternity ? a[8] : 0) + a[32] + a[33]);
        if (clinicType == "S")
        {
            if (claims.FirstOrDefault()?.PlnGpSpShared?.Trim() == "Y") return other + net > annual;
            return ClinicInvoiceRules.NumOrNull(member.SpecBalance) is not null && spec + net > member.SpecAccum;
        }
        return other + net > annual;
    }

    private async Task<AdjustmentCase> Load(IDbSession s, string cas)
    {
        var (row, claim, paid, posted, ids, coPay, email, remark, btbg) = await repository.LoadAsync(s, cas);
        if (row is null) throw new NotFoundException(NotFound);
        var bordxed = T(row, "INVBordxed") is "1" or "True";
        var (mode, msg) = posted && !paid && !bordxed ? (AdjustmentEditMode.DiagnosisOnly, PostedOnly)
            : paid && bordxed && posted ? (AdjustmentEditMode.DiagnosisOnly, PostedOnly)
            : bordxed ? (AdjustmentEditMode.None, Bordered)
            : paid ? (AdjustmentEditMode.None, PaidDone)
            : (AdjustmentEditMode.Full, (string?)null);
        var numericIds = ids.Count > 0 && string.Join("", ids).Replace(",", "", StringComparison.Ordinal).All(char.IsDigit);
        var reasonIds = numericIds ? ids.SelectMany(i => i.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)).ToList() : [];
        var reasons = reasonIds.Count > 0 ? [] : await repository.PendingReasonsAsync(s, T(row, "INVPayTo") is { Length: > 0 } p ? p[..1] : "C");
        var messages = new List<string>();
        if (await repository.SameVisitCasesAsync(s, cas) is { Count: >= 2 } same)
            messages.Add("Multiple cases Found for this Member on same VistDate: " + string.Join("", same.Select(c => c + ",")));
        if (T(row, "AuditQuery").Length > 0) messages.Add(Audited);
        var tables = await repository.WorksheetAsync(s, cas);
        var lookups = await worksheet.LookupsAsync(s, new WorksheetInput { ClinicCode = T(row, "ClinicCode") });
        return new AdjustmentCase
        {
            Case = row, Claim = claim, Paid = paid, Posted = posted, Bordxed = bordxed, EditMode = mode, EditMessage = msg, PendingReasonIds = reasonIds,
            PendingReasons = reasons, SpecialistCoPay = coPay, Email = email, SpecialRemark = remark, Btbg = btbg,
            Worksheet = InvoiceEnquiryService.Stored(lookups.Categories, tables), Messages = messages,
        };
    }

    /// <inheritdoc />
    public async Task<string> AuditAsync(string? casNumber, AuditReviewRequest r)
    {
        var cas = Case(casNumber);
        await using var s = await sessions.OpenAsync();
        var (row, _, _, _, _, _, _, _, _) = await repository.LoadAsync(s, cas);
        if (row is null) throw new NotFoundException(NotFound);
        if (string.IsNullOrWhiteSpace(r.Query)) throw new BusinessRuleException("Please Enter Query! ");
        if (string.IsNullOrWhiteSpace(r.AuditBy)) throw new BusinessRuleException("Please Enter Audit By! ");
        if (string.IsNullOrWhiteSpace(r.AuditDate)) throw new BusinessRuleException("Please Enter Audit Date! ");
        var date = ClinicInvoiceRules.Date(r.AuditDate) ?? throw new BusinessRuleException("Invalid date format!");
        await repository.AuditAsync(s, cas, r, date);
        return "Audit Details updated ! ";
    }

    /// <inheritdoc />
    public async Task ClaimsAppAsync(string? casNumber)
    {
        await using var s = await sessions.OpenAsync();
        await repository.ClaimsAppAsync(s, Case(casNumber));
    }

    /// <inheritdoc />
    public async Task<string> VoidBordxAsync(string? casNumber, string userCode)
    {
        if (string.IsNullOrWhiteSpace(casNumber)) throw new BusinessRuleException("Please Insert Case No..!");
        var cas = Case(casNumber);
        await using var s = await sessions.BeginAsync();
        var loaded = await Load(s, cas);
        if (loaded.Bordxed) throw new BusinessRuleException("This case Already Bordxed..!");
        if (loaded.Paid) throw new BusinessRuleException("This case Already Payment Done..!");
        if (T(loaded.Case, "INVPayTo") != "M") throw new BusinessRuleException("VOID Bordx Only For Reimbursement Cases..!");
        var group = (await invoices.MemberAsync(T(loaded.Case, "MBMnumber"), T(loaded.Case, "CoverID"))).PlanGroupCompany;
        var no = await repository.VoidBordxAsync(s, cas, group, userCode.Trim().ToUpperInvariant(), clock.GetLocalNow().Date);
        await s.CommitAsync();
        return $"Void Bordx Generated! Bordx No: {no} ";
    }

    /// <summary>After Bordx's Edit message.</summary>
    public const string PostedNoChange = "This case is posted into SUN System. Changes are not allowed!";

    /// <inheritdoc />
    public async Task<AdjustmentCase> LoadAfterBordxAsync(string? casNumber) => AfterBordx(await LoadAsync(casNumber));

    /// <inheritdoc />
    public async Task<AdjustmentResult> UpdateAfterBordxAsync(AdjustmentRequest r, string userCode)
    {
        var cas = Case(r.CaseNo);
        var e = r.Entry;
        var x = r.Extra;
        var warnings = new List<string>();
        await using var s = await sessions.BeginAsync();
        var loaded = AfterBordx(await Load(s, cas));
        if (loaded.EditMode == AdjustmentEditMode.None) throw new BusinessRuleException(loaded.EditMessage!);
        var member = await invoices.MemberAsync(T(loaded.Case, "MBMnumber"), T(loaded.Case, "CoverID"));
        var line = await Line(s, r, loaded with { SpecialistCoPay = N(loaded.Claim, "CLMCoPayColl") }, member, userCode);
        var clinic = (e.ClinicCode ?? "").Trim().ToUpperInvariant();
        var clinicName = clinic is "NP" or "NPS" ? "x" : (await clinics.ClinicAsync(s, clinic))?.Name ?? "";
        var otherDiag = e.Diagnoses.Count > 4 && !string.IsNullOrWhiteSpace(e.Diagnoses[4])
            && (await clinics.DiagnosisAsync(s, e.Diagnoses[4]!.Trim()))?.OtherIndicator?.Trim() == "Y";
        var bill = ClinicInvoiceRules.NumOrNull(e.ClinicBillAmount);
        var memo = F(x.MemoClaim);
        string? message = null;
        if (string.IsNullOrWhiteSpace(e.MbmNumber)) message = "Please keyin Member No ! ";
        else if (string.IsNullOrWhiteSpace(e.VisitDate)) message = "Please Enter Visit Date! ";
        else if (clinicName.Length == 0) message = "Please keyin the Clinic code ! ";
        else if (clinic == "NP" && string.IsNullOrWhiteSpace(x.NpClinic)) message = "Please keyin the NP Clinic ! ";
        else if (clinic == "NPS" && string.IsNullOrWhiteSpace(x.NpClinic)) message = "Please keyin the NPS Clinic ! ";
        else if (otherDiag && string.IsNullOrWhiteSpace(e.DiagnosisOther)) message = "Please Keyin Others Diagnosis ! ";
        else if (F(x.PayTo) == "C" && clinic == "NP") message = "Invalid Clinic Code";
        else if (bill is null && memo == "Y") message = "Please Enter Bill Amount ! ";
        else if (bill == 0 && memo == "Y") message = "Bill Amount Amount cannot be Zero ! ";
        else if ((bill ?? 0) != line.Total && (bill ?? 0) != 0) message = "Please check Bill Amount! ";
        else if (clinic == "NPS" && string.IsNullOrWhiteSpace(x.RefLetter)) message = "Please enter status for referral letter";
        if (message is not null) throw new BusinessRuleException(message);
        // The desktop only warns about a negative total and saves anyway.
        if (line.Total < 0) warnings.Add("Total Invoice Amount cannot be Zero ! ");
        await repository.SaveAsync(s, line with { Request = r, UserCode = userCode.Trim().ToUpperInvariant(), RepAction = T(loaded.Case, "CASRepAction"), AfterBordx = true });
        await s.CommitAsync();
        return new AdjustmentResult("Record updated ! ", T(loaded.Case, "OPBatchNo"), warnings);
    }

    /// <summary>After Posting's Create message.</summary>
    public const string NotPosted = "This case is not yet POST into SUN System. Changes are not allowed!";

    /// <inheritdoc />
    public async Task<AdjustmentCase> LoadAfterPostAsync(string? casNumber) => AfterPost(await LoadAsync(casNumber));

    /// <inheritdoc />
    public async Task<AdjustmentResult> CreateAfterPostAsync(AdjustmentRequest r, string userCode)
    {
        var cas = Case(r.CaseNo);
        var e = r.Entry;
        var x = r.Extra;
        string newCase;
        bool aso;
        await using (var s = await sessions.BeginAsync())
        {
            var loaded = AfterPost(await Load(s, cas));
            if (loaded.EditMode == AdjustmentEditMode.None) throw new BusinessRuleException(loaded.EditMessage!);
            var member = await invoices.MemberAsync(T(loaded.Case, "MBMnumber"), T(loaded.Case, "CoverID"));
            var line = await Line(s, r, loaded, member, userCode);
            var clinic = (e.ClinicCode ?? "").Trim().ToUpperInvariant();
            var clinicName = clinic is "NP" or "NPS" ? "x" : (await clinics.ClinicAsync(s, clinic))?.Name ?? "";
            var otherDiag = e.Diagnoses.Count > 4 && !string.IsNullOrWhiteSpace(e.Diagnoses[4])
                && (await clinics.DiagnosisAsync(s, e.Diagnoses[4]!.Trim()))?.OtherIndicator?.Trim() == "Y";
            var bill = ClinicInvoiceRules.NumOrNull(e.ClinicBillAmount);
            string? message = null;
            if (string.IsNullOrWhiteSpace(e.MbmNumber)) message = "Please keyin Member No ! ";
            else if (string.IsNullOrWhiteSpace(e.VisitDate)) message = "Please Enter Visit Date! ";
            else if (clinicName.Length == 0) message = "Please keyin the Clinic code ! ";
            else if (clinic == "NP" && string.IsNullOrWhiteSpace(x.NpClinic)) message = "Please keyin the NP Clinic ! ";
            else if (clinic == "NPS" && string.IsNullOrWhiteSpace(x.NpClinic)) message = "Please keyin the NPS Clinic ! ";
            else if (otherDiag && string.IsNullOrWhiteSpace(e.DiagnosisOther)) message = "Please Keyin Others Diagnosis ! ";
            else if (F(x.PayTo) == "C" && clinic == "NP") message = "Invalid Clinic Code";
            else if (bill is null && F(x.MemoClaim) == "Y") message = "Please Enter Bill Amount ! ";
            else if ((bill ?? 0) != line.Total) message = "Please check Bill Amount! ";
            else if (clinic == "NPS" && string.IsNullOrWhiteSpace(x.RefLetter)) message = "Please enter status for referral letter";
            else if (line.Total == 0) message = "Total Invoice Amount cannot be Zero ! ";
            if (message is not null) throw new BusinessRuleException(message);
            newCase = await clinics.NextCaseNoAsync(s);
            await repository.SaveAsync(s, line with
            {
                Request = r, Loaded = line.Loaded, UserCode = userCode.Trim().ToUpperInvariant(), RepAction = T(loaded.Case, "CASRepAction"), NewCase = newCase,
            });
            await s.CommitAsync();
            aso = member.AsoInd == "Y";
        }
        if (aso)
        {
            await using var s2 = await sessions.OpenAsync();
            await clinics.AsoUtilisationAsync(s2, newCase);
        }
        return new AdjustmentResult("Record updated ! ", "", [], newCase);
    }

    /// <summary>After Posting's rule: a sub-case can only be created for a case posted to SUN.</summary>
    private static AdjustmentCase AfterPost(AdjustmentCase c) => c with
    {
        EditMode = c.Posted ? AdjustmentEditMode.Full : AdjustmentEditMode.None,
        EditMessage = c.Posted ? null : NotPosted,
    };

    /// <summary>After Bordx's Edit rule: everything unless posted to SUN; its co-pay is the claim's CLMCoPayColl.</summary>
    private static AdjustmentCase AfterBordx(AdjustmentCase c) => c with
    {
        EditMode = c.Posted ? AdjustmentEditMode.None : AdjustmentEditMode.Full,
        EditMessage = c.Posted ? PostedNoChange : null,
        SpecialistCoPay = N(c.Claim, "CLMCoPayColl"),
    };

    private static string Case(string? casNumber)
    {
        var cas = (casNumber ?? "").Trim().ToUpperInvariant();
        return cas.Length is > 0 and <= 20 && cas.All(char.IsLetterOrDigit) ? cas : throw new NotFoundException(NotFound);
    }

    private static string F(string? v) => string.IsNullOrWhiteSpace(v) ? "" : v.Trim()[..1].ToUpperInvariant();

    /// <summary>A trimmed text value of a row.</summary>
    public static string T(Dictionary<string, object?>? r, string k) =>
        r is not null && r.TryGetValue(k, out var v) && v is not null and not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    /// <summary>A number of a row (else 0).</summary>
    public static decimal N(Dictionary<string, object?>? r, string k) => ClinicInvoiceRules.Num(T(r, k));
}
