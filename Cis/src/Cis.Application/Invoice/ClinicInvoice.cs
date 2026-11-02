using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Membership;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Invoice;

/// <summary>ComboListing.</summary>
/// <param name="ReceiptModes">InvRecMode.</param>
/// <param name="ProcessMonths">INVProcMonth (MM/YYYY).</param>
public record ClinicSetup(List<LookupOption> ReceiptModes, List<string> ProcessMonths);

/// <summary>TxtClinicCode_LostFocus.</summary>
/// <param name="Code">Code.</param>
/// <param name="Name">Name (panel clinics only).</param>
/// <param name="ClinicType">G / S (blank → G).</param>
/// <param name="FfsStatus">CLNFFSStatus.</param>
/// <param name="Panel">Panel C / Y.</param>
/// <param name="McoProcFeePercent">CLNMCOProcfeePerc.</param>
public record ClinicInfo(string Code, string Name, string ClinicType, string FfsStatus, bool Panel, decimal? McoProcFeePercent);

/// <summary>VerifyDiagnosis.</summary>
/// <param name="Code">Code.</param>
/// <param name="Description">Description.</param>
/// <param name="OtherIndicator">ReasonOtherDiag (Y = an "others" text is required for diagnosis 5).</param>
public record InvoiceDiagnosis(string Code, string? Description, string? OtherIndicator);

/// <summary>CmdGo look-ups after MemberSuppInfo (cis.InvClinic_MemberExtra).</summary>
/// <param name="Visit">AnnualLimit + PLN visit limits.</param>
/// <param name="Plan">PLN maternity / dental / ASO indicators.</param>
/// <param name="GpSp">GP / SP limits and group company.</param>
/// <param name="ClaimInfo">MBMOthers + GRPCompany.</param>
/// <param name="Aso">Tbl_ASOMemberDetails rows.</param>
public record InvMemberExtra(Dictionary<string, object?>? Visit, Dictionary<string, object?>? Plan, Dictionary<string, object?>? GpSp,
    Dictionary<string, object?>? ClaimInfo, List<Dictionary<string, object?>> Aso);

/// <summary>The member panel of frmInvoice after Go.</summary>
public sealed record InvoiceMember
{
    /// <summary>MemberSuppInfo.</summary>
    public MemberSupp Member { get; init; } = new();
    /// <summary>P principal / S supplementary.</summary>
    public string MemberType { get; init; } = "P";
    /// <summary>Ann. Lmt (amount or FULL).</summary>
    public string AnnualLimit { get; init; } = "";
    /// <summary>Vst Lmt.</summary>
    public string VisitLimit { get; init; } = "";
    /// <summary>Bal. Lmt ("" when the annual limit is not an amount).</summary>
    public string BalanceLimit { get; init; } = "";
    /// <summary>Acc Amt.</summary>
    public decimal AccumAmount { get; init; }
    /// <summary>SP / GP limits.</summary>
    public decimal GpLimit { get; init; }
    /// <summary>SP limit.</summary>
    public decimal SpLimit { get; init; }
    /// <summary>Specialist limit / utilised / balance.</summary>
    public string SpecLimit { get; init; } = "";
    /// <summary>Specialist utilised.</summary>
    public decimal SpecAccum { get; init; }
    /// <summary>Specialist balance.</summary>
    public string SpecBalance { get; init; } = "";
    /// <summary>Top-ups.</summary>
    public decimal Topup { get; init; }
    /// <summary>Specialist top-ups.</summary>
    public decimal SpTopup { get; init; }
    /// <summary>PLN indicators.</summary>
    public string MaternityInd { get; init; } = "";
    /// <summary>Dental.</summary>
    public string DentalInd { get; init; } = "";
    /// <summary>ASO.</summary>
    public string AsoInd { get; init; } = "N";
    /// <summary>PLN.GRPCompany (C014 = Prince Court).</summary>
    public string PlanGroupCompany { get; init; } = "";
    /// <summary>Co-pay (A29 default).</summary>
    public decimal CoPay { get; init; }
    /// <summary>MC days exceeded (MC status Y and no days left).</summary>
    public bool McDaysExceeded { get; init; }
    /// <summary>Look-up warnings.</summary>
    public List<string> Warnings { get; init; } = [];
    /// <summary>Claim cost centre, department code, subsidiary, group company code and department (SaveINVClaim).</summary>
    public string CostCenter { get; init; } = "NA";
    /// <summary>Department code.</summary>
    public string DeptCode { get; init; } = "NA";
    /// <summary>Subsidiary.</summary>
    public string Subsidiary { get; init; } = "NA";
    /// <summary>Group company code.</summary>
    public string GroupCompanyCode { get; init; } = "";
    /// <summary>Department name.</summary>
    public string DeptName { get; init; } = "NA";
}

/// <summary>A computed invoice line (AssignData2ListView + the worksheet).</summary>
/// <param name="Amounts">A1..A33.</param>
/// <param name="Total">IntInvAmt.</param>
/// <param name="ClinicBill">lblCLNBillAmt (total less the co-pay).</param>
/// <param name="ActCoPay">lblActCoPay (row 43).</param>
/// <param name="OtherDiagnosis">Diagnosis 5 needs the "others" text.</param>
/// <param name="Member">Member panel.</param>
/// <param name="Clinic">Clinic.</param>
/// <param name="Worksheet">The worksheet.</param>
public record ClinicInvoiceLine(decimal[] Amounts, decimal Total, decimal ClinicBill, decimal ActCoPay, bool OtherDiagnosis, InvoiceMember Member,
    ClinicInfo Clinic, WorksheetResult Worksheet);

/// <summary>Add result: the line, or the duplicate question to confirm.</summary>
/// <param name="Line">The line.</param>
/// <param name="Confirm">"Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel".</param>
public record ClinicAddResult(ClinicInvoiceLine? Line, string? Confirm);

/// <summary>Add / Verify / Save request.</summary>
/// <param name="Entry">The case.</param>
/// <param name="ConfirmDuplicate">The user answered Yes to the duplicate question.</param>
public record ClinicInvoiceRequest(ClinicInvoiceEntry Entry, bool ConfirmDuplicate = false);

/// <summary>Everything one save writes (cmdSave_Click's routines).</summary>
public sealed record ClinicCaseSave
{
    /// <summary>The case and its line.</summary>
    public string CaseNo { get; init; } = "";
    /// <summary>The entry.</summary>
    public ClinicInvoiceEntry Entry { get; init; } = new();
    /// <summary>The line.</summary>
    public ClinicInvoiceLine Line { get; init; } = null!;
    /// <summary>Member number without the cover.</summary>
    public string MbmNumber { get; init; } = "";
    /// <summary>Cover.</summary>
    public string CoverId { get; init; } = "00";
    /// <summary>User and today.</summary>
    public string UserCode { get; init; } = "";
    /// <summary>Today.</summary>
    public DateTime Today { get; init; }
    /// <summary>SaveINVClaim's computed amounts.</summary>
    public decimal TotalActual { get; init; }
    /// <summary>Final approved.</summary>
    public decimal FinalApprove { get; init; }
    /// <summary>Payable by Medix to the clinic.</summary>
    public decimal Payable { get; init; }
    /// <summary>MCO processing fee.</summary>
    public decimal McoProcFee { get; init; }
    /// <summary>Bordereaux amount.</summary>
    public decimal BordxAmt { get; init; }
    /// <summary>UpdateMBMBalance: limit row (null = the principal's MBMCLNOthers), deduction, FULL.</summary>
    public string? LimitCover { get; init; }
    /// <summary>Balance deduction (null = balance unchanged).</summary>
    public decimal? BalanceDeduction { get; init; }
}

/// <summary>Data access of the clinic invoice (<c>cis.InvClinic_*</c> and the desktop's save procedures).</summary>
public interface IClinicInvoiceRepository
{
    /// <summary>ComboListing.</summary>
    Task<ClinicSetup> SetupAsync(IDbSession session);

    /// <summary>CLN.</summary>
    Task<ClinicInfo?> ClinicAsync(IDbSession session, string code);

    /// <summary>TreatmentReason.</summary>
    Task<InvoiceDiagnosis?> DiagnosisAsync(IDbSession session, string code);

    /// <summary>Go look-ups.</summary>
    Task<InvMemberExtra> MemberExtraAsync(IDbSession session, string mbmNumber, string covId, string planCode);

    /// <summary>CIS_FFS2013 rows.</summary>
    Task<List<InvClaim>> ClaimsAsync(IDbSession session, string mbmNumber, int mode, string covId);

    /// <summary>Same-day cases at the clinic and DiagNotPayable codes for <paramref name="diagnosisPlan"/>.</summary>
    Task<(int SameDayCases, List<string> NotPayable)> ChecksAsync(IDbSession session, string mbmNumber, string covId, DateTime visit, string clinic,
        string diagnosisPlan, IReadOnlyList<string?> diagnoses);

    /// <summary>GenerateCISCaseNo.</summary>
    Task<string> NextCaseNoAsync(IDbSession session);

    /// <summary>SaveCase … SaveDup_ori in the open transaction.</summary>
    Task SaveAsync(IDbSession session, ClinicCaseSave save);

    /// <summary>UpdateASOLimit (sp_ASOClaimsUtilization, after the commit).</summary>
    Task AsoUtilisationAsync(IDbSession session, string caseNo);
}

/// <summary>Invoice > Clinic > Clinic (frmInvoice).</summary>
public interface IClinicInvoiceService
{
    /// <summary>ComboListing.</summary>
    Task<ClinicSetup> SetupAsync();

    /// <summary>TxtClinicCode_LostFocus.</summary>
    Task<ClinicInfo> ClinicAsync(string? code);

    /// <summary>VerifyDiagnosis.</summary>
    Task<InvoiceDiagnosis> DiagnosisAsync(string? code);

    /// <summary>CmdGo_Click.</summary>
    Task<InvoiceMember> MemberAsync(string? mbmNumber, string? covId);

    /// <summary>CmdAddRec_Click (ChkData, AssignData2ListView, the worksheet).</summary>
    Task<ClinicAddResult> AddAsync(ClinicInvoiceRequest request, string userCode);

    /// <summary>cmdVerify_Click: the errors (empty = "Verifed Sucessfully...").</summary>
    Task<List<string>> VerifyAsync(ClinicInvoiceRequest request, string userCode);

    /// <summary>cmdSave_Click: the new case number.</summary>
    Task<string> SaveAsync(ClinicInvoiceRequest request, string userCode);
}

/// <inheritdoc />
public sealed class ClinicInvoiceService(IDbSessions sessions, IClinicInvoiceRepository repository, IMemberSuppRepository members,
    IMbmAdjustmentRepository adjustment, IMbmEnquiryRepository enquiry, IWorksheetRepository worksheet, TimeProvider clock) : IClinicInvoiceService
{
    /// <summary>Desktop texts.</summary>
    public const string MemberFirst = "Please Key Member No Or IC Number ! ", NoClinic = "No such Clinic Code ! ", NonPanel = "Non-Panel Clinic !",
        InvalidDiagnosis = "Invalid Diagnosis code ! Please try again ", Duplicate = "Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel",
        McExceeded = "MC Days Exceeded! ", NoBalance = "No Balance Limit to Register Case. Please Check!!!", Saved = "Record Saved ! ",
        Verified = "Verifed Sucessfully...";

    /// <inheritdoc />
    public async Task<ClinicSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<ClinicInfo> ClinicAsync(string? code)
    {
        await using var s = await sessions.OpenAsync();
        return await Clinic(s, code);
    }

    /// <inheritdoc />
    public async Task<InvoiceDiagnosis> DiagnosisAsync(string? code)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.DiagnosisAsync(s, (code ?? "").Trim()) ?? throw new NotFoundException(InvalidDiagnosis);
    }

    /// <inheritdoc />
    public async Task<InvoiceMember> MemberAsync(string? mbmNumber, string? covId)
    {
        await using var s = await sessions.OpenAsync();
        return await Member(s, mbmNumber, covId);
    }

    /// <inheritdoc />
    public async Task<ClinicAddResult> AddAsync(ClinicInvoiceRequest r, string userCode)
    {
        await using var s = await sessions.OpenAsync();
        var (line, sameDay, _) = await Build(s, r.Entry, userCode);
        return sameDay > 0 && !r.ConfirmDuplicate ? new ClinicAddResult(null, Duplicate) : new ClinicAddResult(line, null);
    }

    /// <inheritdoc />
    public async Task<List<string>> VerifyAsync(ClinicInvoiceRequest r, string userCode)
    {
        await using var s = await sessions.OpenAsync();
        var (line, _, notPayable) = await Build(s, r.Entry, userCode);
        return Errors(r.Entry, line, notPayable);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(ClinicInvoiceRequest r, string userCode)
    {
        string caseNo;
        bool aso;
        await using (var s = await sessions.BeginAsync())
        {
            var (line, sameDay, _) = await Build(s, r.Entry, userCode);
            if (sameDay > 0 && !r.ConfirmDuplicate) throw new BusinessRuleException(Duplicate);
            // The desktop enables Save after Verify even when it listed errors (they are only exported).
            var a = line.Amounts;
            var others = a[32] + a[33];
            var clnBill = ClinicInvoiceRules.Num(r.Entry.ClinicBillAmount);
            var limitOk = line.Clinic.ClinicType == "G" && ClinicInvoiceRules.NumOrNull(line.Member.BalanceLimit) is { } bal ? bal - clnBill + others >= 0
                : line.Clinic.ClinicType == "S" && ClinicInvoiceRules.NumOrNull(line.Member.SpecBalance) is { } sp ? sp - clnBill + others >= 0
                : true;
            if (!limitOk) throw new BusinessRuleException(NoBalance);
            caseNo = await repository.NextCaseNoAsync(s);
            await repository.SaveAsync(s, SaveModel(caseNo, r.Entry, line, userCode, clock.GetLocalNow().Date));
            await s.CommitAsync();
            aso = line.Member.AsoInd == "Y";
        }
        if (aso)
        {
            await using var s2 = await sessions.OpenAsync();
            await repository.AsoUtilisationAsync(s2, caseNo);
        }
        return caseNo;
    }

    /// <summary>cmdVerify_Click's errors (the plan used for DiagNotPayable is the desktop's row 64, i.e. diagnosis 3 — reproduced).</summary>
    private List<string> Errors(ClinicInvoiceEntry e, ClinicInvoiceLine line, List<string> notPayable) =>
        ClinicInvoiceRules.Verify(e, line.Total, line.Member.Member.EffDate, line.Member.Member.ExpDate, line.Member.BalanceLimit, line.Member.VisitLimit,
            notPayable, clock.GetLocalNow().Date);

    /// <summary>The line, the same-day case count and the not-payable diagnoses.</summary>
    private async Task<(ClinicInvoiceLine Line, int SameDay, List<string> NotPayable)> Build(IDbSession s, ClinicInvoiceEntry e, string userCode)
    {
        var today = clock.GetLocalNow().Date;
        var member = await Member(s, e.MbmNumber, e.CoverId);
        var m = member.Member;
        var clinic = await Clinic(s, e.ClinicCode);
        if (ClinicInvoiceRules.Date(e.VisitDate) is { } visit
            && ClinicInvoiceRules.VisitCheck(visit, today, m.CoverId == "00", m.PolicyStatus, m.ExpDate, m.CancelDate, m.CoverCancelDate, m.CoverClinicExpDate)
                is { } visitMessage)
            throw new BusinessRuleException(visitMessage);
        var d5 = e.Diagnoses.Count > 4 ? e.Diagnoses[4]?.Trim() : null;
        var other = !string.IsNullOrEmpty(d5) && (await repository.DiagnosisAsync(s, d5))?.OtherIndicator?.Trim() == "Y";
        var a = ClinicInvoiceRules.Amounts(e);
        // DisplayBillAmt: Prince Court (C014) charges a co-pay of 5 when a GP fee is entered.
        if (member.PlanGroupCompany == "C014") a[29] = Enumerable.Range(1, 4).Any(i => i < e.B.Count && e.B[i] is not null) ? 5 : 0;
        var total = ClinicInvoiceRules.Total(a);
        if (ClinicInvoiceRules.Check(e, a, total, m.EffDate, m.ExpDate, other, today) is { } message) throw new BusinessRuleException(message);
        var actCoPay = member.PlanGroupCompany == "C014" ? a[29] : member.CoPay;
        var input = new WorksheetInput
        {
            Caller = WorksheetCaller.ClinicInvoice, PayTo = "C", ClinicCode = clinic.Code, ClinicBenefitCard = "", PlanCode = m.PlanCode, InsCode = m.InsType,
            HltCode = m.HltCode, Items = ClinicInvoiceRules.GridItems(a), NonGpAmount = a[31], Row70BalanceLimit = member.BalanceLimit,
            BalanceLimit = member.BalanceLimit, VisitLimit = member.VisitLimit, AnnualIndicator = m.AnnualIndicator != "N", VisitIndicator = m.VisitIndicator != "N",
            CoPayCollected = actCoPay, CoPay = 0, AdjustedAnnualLimit = e.AdjustedAnnualLimit ?? 0,
            GpStatus = Enumerable.Range(1, 4).Sum(i => i < e.B.Count ? e.B[i] ?? 0 : 0) > 0, McoProcStatus = e.Ffs?.Trim().StartsWith('Y') == true,
        };
        var ws = WorksheetEngine.Calculate(input, await worksheet.LookupsAsync(s, input), today, userCode);
        var line = new ClinicInvoiceLine(a, total, total - a[29], actCoPay, other, member, clinic, ws);
        var (sameDay, notPayable) = await repository.ChecksAsync(s, m.MbmNumber, m.CoverId, ClinicInvoiceRules.Date(e.VisitDate)!.Value, clinic.Code,
            e.Diagnoses.Count > 2 ? e.Diagnoses[2]?.Trim() ?? "" : "", e.Diagnoses);
        return (line, sameDay, notPayable);
    }

    /// <summary>SaveINVClaim's amounts and UpdateMBMBalance's row.</summary>
    public static ClinicCaseSave SaveModel(string caseNo, ClinicInvoiceEntry e, ClinicInvoiceLine line, string userCode, DateTime today)
    {
        var v = line.Worksheet.RowValues;
        decimal R(int i) => v.TryGetValue(i, out var x) && x is decimal d ? d : 0;
        var totalActual = R(160) != 0 ? R(160) : line.ClinicBill;
        var finalApprove = R(161) > 0 ? R(161) : totalActual;
        if (finalApprove == 0) finalApprove = line.ClinicBill;
        var mco = e.Ffs?.Trim().StartsWith('Y') == true ? (line.Clinic.McoProcFeePercent ?? 0) * totalActual / 100 : 0;
        var payable = R(158) > 0 ? R(158) : totalActual - mco;
        var bordx = totalActual + line.Amounts[31];
        if (bordx == 0) bordx = line.ClinicBill;
        var m = line.Member.Member;
        var limitCover = m.CoverId == "00" ? null : m.SuppLimitType is "B" or "D" ? "01" : m.SuppLimitType == "C" ? m.CoverId : null;
        var noDed = line.Amounts[32] + line.Amounts[33] + line.Amounts[8];
        decimal? deduction = m.GrossStatus switch { "Y" => R(158) - noDed, "N" => R(161) - noDed, _ => null };
        return new ClinicCaseSave
        {
            CaseNo = caseNo, Entry = e, Line = line, MbmNumber = m.MbmNumber, CoverId = m.CoverId, UserCode = userCode.Trim().ToUpperInvariant(), Today = today,
            TotalActual = totalActual, FinalApprove = finalApprove, Payable = payable, McoProcFee = mco, BordxAmt = bordx, LimitCover = limitCover,
            BalanceDeduction = line.Member.BalanceLimit == "FULL" ? null : deduction,
        };
    }

    private async Task<ClinicInfo> Clinic(IDbSession s, string? code)
    {
        var c = (code ?? "").Trim().ToUpperInvariant();
        if (c.Length == 0) throw new BusinessRuleException("Please keyin the Clinic code ! ");
        var clinic = await repository.ClinicAsync(s, c) ?? throw new NotFoundException(NoClinic);
        return clinic.Panel ? clinic : throw new BusinessRuleException(NonPanel);
    }

    /// <summary>CmdGo_Click: MemberSuppInfo plus the limits, top-ups and utilisation.</summary>
    private async Task<InvoiceMember> Member(IDbSession s, string? mbmNumber, string? covId)
    {
        var text = (mbmNumber ?? "").Trim().ToUpperInvariant();
        if (text.Length == 0) throw new BusinessRuleException(MemberFirst);
        var dash = text.IndexOf('-');
        var number = dash > 0 ? text[..dash].Trim() : text;
        var cov = !string.IsNullOrWhiteSpace(covId) ? covId.Trim() : dash > 0 ? text[(dash + 1)..].Trim() : "00";
        var m = MemberSuppService.Build(number, cov, await members.RowsAsync(s, number, cov));
        var x = await repository.MemberExtraAsync(s, number, cov, m.PlanCode);
        var type = cov == "00" ? "P" : "S";
        var warnings = new List<string>();
        string T(Dictionary<string, object?>? r, string k) => r?.GetValueOrDefault(k) is { } o ? Convert.ToString(o, CultureInfo.InvariantCulture)!.Trim() : "";
        decimal N(Dictionary<string, object?>? r, string k) => ClinicInvoiceRules.Num(T(r, k));
        var aso = T(x.Plan, "AsoInd") is { Length: > 0 } ai ? ai : "N";
        bool maternity = T(x.Plan, "MaternityInd") == "Y", dental = T(x.Plan, "DentalInd") == "Y";

        string annual, visit = "", balance = "";
        decimal topup = 0, spTopup = 0, asoUtilised = 0;
        if (aso == "Y")
        {
            var row = x.Aso.LastOrDefault();
            var combined = T(row, "IsIPOPCombined") == "Y";
            annual = Money(combined ? N(row, "MBMASOLimit") : N(row, "MBMOPLimit"));
            asoUtilised = combined ? N(row, "MBMIPUtilized") + N(row, "MBMOPUtilized") : N(row, "MBMOPUtilized");
        }
        else
        {
            var limits = m.EffDate is null ? null : await adjustment.LimitsAsync(s, type, m.InsType, m.PlanCode, m.EffDate.Value, m.Payor, m.SuppLimitType);
            if (limits is { AnnualOk: false }) warnings.Add(MbmAdjustmentService.NoAnnual);
            annual = limits is null ? "" : limits.AnnualFull ? "FULL" : Money(limits.AnnualLimit);
            if (await enquiry.TopupAsync(s, number, cov) is { } t) (topup, spTopup) = t;
            var visitLookup = limits is null ? "" : limits.VisitFull ? "FULL" : Money(limits.VisitLimit);
            visit = ClinicInvoiceRules.NumOrNull(visitLookup) is null ? visitLookup
                : ClinicInvoiceRules.NumOrNull(m.VisitLimit) is { } mv ? Money(mv) : m.VisitLimit;
        }

        // CallLimit: the specialist limit of insured type = the member number's third character, health N, payor DC.
        var specLimit = MbmEnquiryService.SpecLimit(await enquiry.SpecLimitAsync(s, m.PlanCode, "DC", number.Length > 2 ? number[2..3] : ""), warnings, type == "S");
        var mode = ClinicInvoiceRules.ClaimMode(m.SuppLimitType, cov);
        var claims = mode is { } md ? await repository.ClaimsAsync(s, number, md, cov) : [];
        var (accum, specAccum) = ClinicInvoiceRules.Utilisation(claims, maternity, dental, specLimit != "N/A");
        if (aso == "Y") accum = asoUtilised;
        var specBalance = ClinicInvoiceRules.NumOrNull(specLimit) is { } sl ? Money(sl + spTopup - specAccum) : "";
        if (m.SuppLimitType == "F" && type == "S" && m.RelationType is "D" or "S" && T(x.Visit, "ChildLimitInd") == "Y") annual = Money(N(x.Visit, "ChildDepLimit"));
        if (ClinicInvoiceRules.NumOrNull(annual) is { } ann) balance = Money(ann + topup - accum);

        var mcExceeded = m.McStatus == "Y" && m.McDaysBalance == 0;
        if (mcExceeded) warnings.Add(McExceeded);
        string Na(string k) => T(x.ClaimInfo, k) is { Length: > 0 } v ? v : "NA";
        return new InvoiceMember
        {
            Member = m, MemberType = type, AnnualLimit = annual, VisitLimit = visit, BalanceLimit = balance, AccumAmount = accum, GpLimit = N(x.GpSp, "GPLimit"),
            SpLimit = N(x.GpSp, "SPLimit"), SpecLimit = specLimit, SpecAccum = specAccum, SpecBalance = specBalance, Topup = topup, SpTopup = spTopup,
            MaternityInd = maternity ? "Y" : T(x.Plan, "MaternityInd"), DentalInd = T(x.Plan, "DentalInd"), AsoInd = aso, PlanGroupCompany = T(x.GpSp, "GRPCompany"),
            CoPay = ClinicInvoiceRules.Num(m.CoPay), McDaysExceeded = mcExceeded, Warnings = warnings, CostCenter = Na("MBMCostCenter"), DeptCode = Na("MBMDeptCode"),
            Subsidiary = Na("Subsidiary"), GroupCompanyCode = T(x.ClaimInfo, "GRPCoCode"), DeptName = Na("MBMDepartment"),
        };
    }

    private static string Money(decimal v) => v.ToString("#,##0.00", CultureInfo.InvariantCulture);
}

/// <summary>Column lengths of the case (CISCAS / ClaimInvoice).</summary>
public sealed class ClinicInvoiceRequestValidator : Common.RequestValidator<ClinicInvoiceRequest>
{
    /// <summary>Defines the rules.</summary>
    public ClinicInvoiceRequestValidator()
    {
        RuleFor(r => r.Entry).NotNull();
        RuleFor(r => r.Entry.InvoiceNo).MaximumLength(50);
        RuleFor(r => r.Entry.BillNo).MaximumLength(50);
        RuleFor(r => r.Entry.MbmNumber).MaximumLength(50);
        RuleFor(r => r.Entry.Remark).MaximumLength(500);
        RuleFor(r => r.Entry.DiagnosisOther).MaximumLength(200);
        RuleFor(r => r.Entry.A.Count).LessThanOrEqualTo(34);
        RuleFor(r => r.Entry.B.Count).LessThanOrEqualTo(9);
        RuleFor(r => r.Entry.Medications.Count).LessThanOrEqualTo(7);
    }
}
