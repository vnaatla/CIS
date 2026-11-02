using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Invoice;

/// <summary>The member invoice's own fields (frmInvoiceMBM).</summary>
public sealed record MemberInvoiceExtra
{
    /// <summary>Non-panel clinic name (clinic NP / NPS).</summary>
    public string? NpClinic { get; init; }
    /// <summary>Clinic type G / S (cboCLNType).</summary>
    public string? ClinicType { get; init; }
    /// <summary>Accepted status (cboMemoClaim) Y / N.</summary>
    public string? MemoClaim { get; init; }
    /// <summary>Pharmacy Y / N (IBM members).</summary>
    public string? Pharmacy { get; init; }
    /// <summary>Referral letter Y / N.</summary>
    public string? RefLetter { get; init; }
    /// <summary>Repudiation reason / action / amount (accepted N).</summary>
    public string? RepReason { get; init; }
    /// <summary>Repudiation action.</summary>
    public string? RepAction { get; init; }
    /// <summary>Repudiation amount.</summary>
    public string? RepAmount { get; init; }
}

/// <summary>Add / Verify / Save request of the member invoice.</summary>
/// <param name="Entry">The case (shared with the clinic invoice).</param>
/// <param name="Extra">The member invoice's own fields.</param>
/// <param name="ConfirmReversal">"Is this Reversal Claim? " answered Yes (negative amount with a same-day case).</param>
public record MemberInvoiceRequest(ClinicInvoiceEntry Entry, MemberInvoiceExtra Extra, bool ConfirmReversal = false);

/// <summary>Lengths.</summary>
public sealed class MemberInvoiceRequestValidator : Common.RequestValidator<MemberInvoiceRequest>
{
    /// <summary>Defines the rules.</summary>
    public MemberInvoiceRequestValidator()
    {
        RuleFor(r => r.Entry).NotNull();
        RuleFor(r => r.Extra).NotNull();
        RuleFor(r => r.Entry.InvoiceNo).MaximumLength(50);
        RuleFor(r => r.Entry.BillNo).MaximumLength(50);
        RuleFor(r => r.Entry.Remark).MaximumLength(500);
        RuleFor(r => r.Extra.NpClinic).MaximumLength(200);
        RuleFor(r => r.Extra.RepReason).MaximumLength(500);
        RuleFor(r => r.Extra.RepAction).MaximumLength(500);
        RuleFor(r => r.Entry.A.Count).LessThanOrEqualTo(34);
        RuleFor(r => r.Entry.B.Count).LessThanOrEqualTo(9);
        RuleFor(r => r.Entry.Medications.Count).LessThanOrEqualTo(7);
    }
}

/// <summary>A member invoice line.</summary>
/// <param name="Amounts">A1..A33 (A29 = the co-pay after the IBM / MM rules).</param>
/// <param name="Total">IntInvAmt (all amounts including A29).</param>
/// <param name="Member">Member panel.</param>
/// <param name="Clinic">Clinic.</param>
/// <param name="OtherDiagnosis">Diagnosis 5 needs the others text.</param>
/// <param name="Ibm">IBM member.</param>
/// <param name="Worksheet">The worksheet.</param>
public record MemberInvoiceLine(decimal[] Amounts, decimal Total, InvoiceMember Member, ClinicInfo Clinic, bool OtherDiagnosis, bool Ibm, WorksheetResult Worksheet);

/// <summary>cmdVerify_Click of the member invoice.</summary>
/// <param name="Errors">Listed errors (exported by the desktop).</param>
/// <param name="Duplicate">"Invoice has been entered . MBMNo = …" when same-day cases exist.</param>
/// <param name="AskReversal">The amount is negative: ask "Is this Reversal Claim? ".</param>
/// <param name="Message">Final message (Please check balance limit / Visit date IS GREATER THAN Expiry date / Verifed Sucessfully...).</param>
/// <param name="CanSave">Save is shown.</param>
public record MemberVerifyResult(List<string> Errors, string? Duplicate, bool AskReversal, string? Message, bool CanSave);

/// <summary>The member invoice save (all values the procedures need).</summary>
/// <param name="CaseNo">Case.</param>
/// <param name="OpBatchNo">OC batch of the save.</param>
/// <param name="Request">The request.</param>
/// <param name="Line">The line.</param>
/// <param name="UserCode">User.</param>
/// <param name="Today">Today.</param>
public record MemberCaseSave(string CaseNo, string OpBatchNo, MemberInvoiceRequest Request, MemberInvoiceLine Line, string UserCode, DateTime Today);

/// <summary>Data access of the member invoice (<c>cis.InvMember_*</c>; Go reuses the clinic invoice).</summary>
public interface IMemberInvoiceRepository
{
    /// <summary>sp_IBMMemberInformation.</summary>
    Task<bool> IsIbmAsync(IDbSession session, string mbmNumber);

    /// <summary>The OC batch (GenerateOPSeq "OC").</summary>
    Task<string> NextBatchAsync(IDbSession session);

    /// <summary>The inserts and the member's balance.</summary>
    Task SaveAsync(IDbSession session, MemberCaseSave save);
}

/// <summary>Invoice > Clinic > Member (frmInvoiceMBM).</summary>
public interface IMemberInvoiceService
{
    /// <summary>TxtClinicCode_LostFocus (NP / NPS = a free-text non-panel clinic).</summary>
    Task<ClinicInfo> ClinicAsync(string? code);

    /// <summary>CmdAddRec_Click.</summary>
    Task<MemberInvoiceLine> AddAsync(MemberInvoiceRequest request, string userCode);

    /// <summary>cmdVerify_Click.</summary>
    Task<MemberVerifyResult> VerifyAsync(MemberInvoiceRequest request, string userCode);

    /// <summary>cmdSave_Click: the case number.</summary>
    Task<string> SaveAsync(MemberInvoiceRequest request, string userCode);
}

/// <inheritdoc />
public sealed class MemberInvoiceService(IDbSessions sessions, IMemberInvoiceRepository repository, IClinicInvoiceRepository clinics,
    IClinicInvoiceService invoices, IWorksheetRepository worksheet, TimeProvider clock) : IMemberInvoiceService
{
    /// <summary>Desktop texts.</summary>
    public const string NonPanel = "Non-Panel Clinic !, Press 'Yes' to Proceed, 'No' to Cancel", OverBalance =
        "Bill Amount exceeded Balance Limit entitlement! Please check the charged Bill Amount should not more than Balance Limit.",
        CheckBalance = "Please check balance limit", Expired = "Visit date IS GREATER THAN Expiry date", Reversal = "Is this Reversal Claim? ";

    /// <inheritdoc />
    public async Task<ClinicInfo> ClinicAsync(string? code)
    {
        await using var s = await sessions.OpenAsync();
        return await Clinic(s, code);
    }

    /// <inheritdoc />
    public async Task<MemberInvoiceLine> AddAsync(MemberInvoiceRequest r, string userCode)
    {
        await using var s = await sessions.OpenAsync();
        return await Build(s, r, userCode);
    }

    /// <inheritdoc />
    public async Task<MemberVerifyResult> VerifyAsync(MemberInvoiceRequest r, string userCode)
    {
        await using var s = await sessions.OpenAsync();
        var line = await Build(s, r, userCode);
        return await Verify(s, r, line);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(MemberInvoiceRequest r, string userCode)
    {
        string caseNo;
        bool aso;
        await using (var s = await sessions.BeginAsync())
        {
            var line = await Build(s, r, userCode);
            var v = await Verify(s, r, line);
            if (!v.CanSave) throw new BusinessRuleException(v.Message ?? v.Duplicate ?? v.Errors.FirstOrDefault() ?? "Please verify the invoice.");
            var total = ClinicInvoiceRules.Num(r.Entry.TotalInvoiceAmount);
            if (ClinicInvoiceRules.NumOrNull(line.Member.BalanceLimit) is { } bal && !((total == 0 && bal == 0) || bal > 0))
                throw new BusinessRuleException(ClinicInvoiceService.NoBalance);
            var batch = await repository.NextBatchAsync(s);
            caseNo = await clinics.NextCaseNoAsync(s);
            await repository.SaveAsync(s, new MemberCaseSave(caseNo, batch, r, line, userCode.Trim().ToUpperInvariant(), clock.GetLocalNow().Date));
            await s.CommitAsync();
            aso = line.Member.AsoInd == "Y";
        }
        if (aso)
        {
            await using var s2 = await sessions.OpenAsync();
            await clinics.AsoUtilisationAsync(s2, caseNo);
        }
        return caseNo;
    }

    /// <summary>ChkData of the member invoice: the first desktop message, or null.</summary>
    public static string? Check(ClinicInvoiceEntry e, MemberInvoiceExtra x, decimal total, bool otherDiagnosis)
    {
        var clinic = (e.ClinicCode ?? "").Trim().ToUpperInvariant();
        var memo = First(x.MemoClaim);
        if (string.IsNullOrWhiteSpace(e.MbmNumber)) return "Please keyin Member No ! ";
        if (clinic.Length == 0) return "Please keyin the Clinic code ! ";
        if (string.IsNullOrWhiteSpace(e.InvoiceNo)) return "Please Enter Invoice No ! ";
        if (string.IsNullOrWhiteSpace(e.InvoiceDate)) return "Please Enter Invoice Date! ";
        if (string.IsNullOrWhiteSpace(e.ReceiptDate)) return "Please Enter Invoice Receipt Date ! ";
        if (string.IsNullOrWhiteSpace(e.ReceiptMode)) return "Please Enter Invoice Receipt Mode ! ";
        if (memo == "Y" && string.IsNullOrWhiteSpace(e.TotalInvoiceAmount)) return "Please Keyin Invoice Amount ! ";
        if (string.IsNullOrWhiteSpace(e.VisitDate)) return "Please Enter Visit Date! ";
        if (memo == "Y" && total == 0) return "Total Invoice Amount cannot be Zero ! ";
        if (clinic is "NP" or "NPS" && string.IsNullOrWhiteSpace(x.NpClinic)) return "Please keyin the NP Clinic ! ";
        if (otherDiagnosis && string.IsNullOrWhiteSpace(e.DiagnosisOther)) return "Please Keyin Others Diagnosis ! ";
        if (clinic == "NP" && string.IsNullOrWhiteSpace(x.ClinicType)) return "Please Enter Clinic type. ";
        if (memo.Length == 0) return "Please Enter Accepted status ";
        if (memo == "N" && string.IsNullOrWhiteSpace(x.RepReason)) return "Please Enter Repudiation Reason";
        if (memo == "N" && string.IsNullOrWhiteSpace(x.RepAction)) return "Please Enter Action";
        if (memo == "N" && string.IsNullOrWhiteSpace(x.RepAmount)) return "Please Enter Action";
        if (clinic == "NPS" && string.IsNullOrWhiteSpace(x.RefLetter)) return "Please enter status for referral letter";
        return null;
    }

    /// <summary>The co-pay A29: ChkData's MM rule (referral letter N, not pharmacy) then isIBMUser.</summary>
    public static decimal CoPay(decimal[] a, MemberInvoiceExtra x, decimal mmPercent, bool ibm, DateTime? visit)
    {
        decimal Sum() => Enumerable.Range(1, 33).Sum(i => a[i]);
        var copay = a[29];
        // strPLNMMPerc is never blank (MemberSuppInfo gives 0), so only the referral letter and pharmacy decide.
        if (First(x.RefLetter) == "N" && First(x.Pharmacy) == "N") copay = -(Sum() * (1 - mmPercent));
        a[29] = copay;
        if (!ibm || visit is null) return copay;
        var total = Sum();
        var type = First(x.ClinicType);
        if (visit.Value.Year >= 2014)
        {
            if (First(x.Pharmacy) == "Y") copay = -(total / 10);
            else if (type == "G") copay = First(x.MemoClaim) == "Y" ? -5 : 0;
            else if (type == "S") copay = -(total / 4);
        }
        else if (visit.Value.Year == 2013 && type == "S" && First(x.RefLetter) == "N") copay = -(total / 4);
        return Math.Round(copay, 2, MidpointRounding.AwayFromZero);
    }

    private async Task<ClinicInfo> Clinic(IDbSession s, string? code)
    {
        var c = (code ?? "").Trim().ToUpperInvariant();
        if (c.Length == 0) throw new BusinessRuleException("Please keyin the Clinic code ! ");
        var clinic = await clinics.ClinicAsync(s, c) ?? throw new NotFoundException(ClinicInvoiceService.NoClinic);
        // NP / NPS are panel entries whose clinic name is typed; any other non-panel clinic must be registered first (Entity maintenance).
        return clinic.Panel ? clinic : throw new BusinessRuleException(NonPanel);
    }

    private async Task<MemberInvoiceLine> Build(IDbSession s, MemberInvoiceRequest r, string userCode)
    {
        var e = r.Entry;
        var x = r.Extra ?? new MemberInvoiceExtra();
        var today = clock.GetLocalNow().Date;
        var member = await invoices.MemberAsync(e.MbmNumber, e.CoverId);
        var m = member.Member;
        var clinic = await Clinic(s, e.ClinicCode);
        var visit = ClinicInvoiceRules.Date(e.VisitDate);
        if (visit is { } v && ClinicInvoiceRules.VisitCheck(v, today, m.CoverId == "00", m.PolicyStatus, m.ExpDate, m.CancelDate, m.CoverCancelDate,
                m.CoverClinicExpDate) is { } visitMessage)
            throw new BusinessRuleException(visitMessage);
        if (ClinicInvoiceRules.NumOrNull(member.BalanceLimit) is { } bal && bal - ClinicInvoiceRules.Num(e.TotalInvoiceAmount) < 0)
            throw new BusinessRuleException(OverBalance);
        var d5 = e.Diagnoses.Count > 4 ? e.Diagnoses[4]?.Trim() : null;
        var other = !string.IsNullOrEmpty(d5) && (await clinics.DiagnosisAsync(s, d5))?.OtherIndicator?.Trim() == "Y";
        var a = ClinicInvoiceRules.Amounts(e);
        if (Check(e, x, Enumerable.Range(1, 33).Sum(i => a[i]), other) is { } message) throw new BusinessRuleException(message);
        var ibm = await repository.IsIbmAsync(s, m.MbmNumber);
        a[29] = CoPay(a, x, ClinicInvoiceRules.Num(m.MmPercent), ibm, visit);
        var total = Math.Round(Enumerable.Range(1, 33).Sum(i => a[i]), 2, MidpointRounding.AwayFromZero);
        var input = new WorksheetInput
        {
            Caller = WorksheetCaller.MemberInvoice, PayTo = "M", ClinicCode = clinic.Code, PlanCode = m.PlanCode, InsCode = m.InsType, HltCode = m.HltCode,
            Items = ClinicInvoiceRules.GridItems(a), NonGpAmount = a[31], Row70BalanceLimit = member.BalanceLimit, BalanceLimit = member.BalanceLimit,
            VisitLimit = member.VisitLimit, AnnualIndicator = m.AnnualIndicator != "N", VisitIndicator = m.VisitIndicator != "N", CoPayCollected = a[29],
            CoPay = 0, AdjustedAnnualLimit = e.AdjustedAnnualLimit ?? 0, ApprovedOverride = m.BordxFees,
        };
        var ws = WorksheetEngine.Calculate(input, await worksheet.LookupsAsync(s, input), today, userCode);
        return new MemberInvoiceLine(a, total, member, clinic, other, ibm, ws);
    }

    /// <summary>cmdVerify_Click's rules.</summary>
    private async Task<MemberVerifyResult> Verify(IDbSession s, MemberInvoiceRequest r, MemberInvoiceLine line)
    {
        var e = r.Entry;
        var m = line.Member.Member;
        var today = clock.GetLocalNow().Date;
        var visit = ClinicInvoiceRules.Date(e.VisitDate);
        var (sameDay, notPayable) = await clinics.ChecksAsync(s, m.MbmNumber, m.CoverId, visit ?? today, line.Clinic.Code,
            e.Diagnoses.Count > 2 ? e.Diagnoses[2]?.Trim() ?? "" : "", e.Diagnoses);
        var errors = notPayable.Select(d => $"Diagnosis '{d}' is not payable ! ").ToList();
        if (string.IsNullOrWhiteSpace(e.InvoiceNo)) errors.Add("Invoice No Can't be null ! ");
        DateTime? eff = m.EffDate, exp = m.ExpDate;
        if (ClinicInvoiceRules.Date(e.InvoiceDate) is { } inv && eff is not null && exp is not null)
        {
            if (inv < eff) errors.Add("Invoice Date MUST BE GREATER than Clinical Effective date ! ");
            if (inv > exp.Value.AddDays(14)) errors.Add("Invoice date IS GREATER THAN Clinical Expiry date + 14 days ! ");
            if ((today - inv).TotalDays > 45) errors.Add("45 Days LATE ! ");
        }
        bool expired = false, overLimit = false;
        if (visit is { } vd && eff is not null && exp is not null)
        {
            if (vd < eff) errors.Add("Visit date MUST BE AFTER the Effective date ! ");
            if (vd > exp.Value.AddDays(14)) errors.Add("Visit date IS GREATER THAN Expiry date + 14 days ! ");
            if (vd > exp)
            {
                expired = true;
                errors.Add("");
            }
            else
            {
                if (ClinicInvoiceRules.NumOrNull(line.Member.BalanceLimit) is { } bal && line.Total > bal + line.Amounts[33])
                {
                    errors.Add("Invoice Amt > Balance Limit ! ");
                    overLimit = true;
                }
                if (ClinicInvoiceRules.NumOrNull(line.Member.VisitLimit) is { } vl && line.Total > vl) errors.Add("Invoice Amt > Visit Limit ! ");
            }
        }
        var duplicate = sameDay > 0 ? $"Invoice has been entered . MBMNo ={m.MbmNumber}-{m.CoverId} , Date Visit = {visit:dd/MM/yyyy}" : null;
        var dupBlocks = duplicate is not null && !(line.Total < 0 && r.ConfirmReversal);
        // fval = TxtTotInvoiceAmt + A29 against the line amount (which includes A29).
        var totalsOk = Math.Round(line.Total, 4) == Math.Round(ClinicInvoiceRules.Num(e.TotalInvoiceAmount) + line.Amounts[29], 4);
        if (!totalsOk) errors.Add("Invoice Amt is different from Total Bill Amt");
        string? message = null;
        var canSave = totalsOk;
        if (dupBlocks) canSave = false;
        else if (overLimit) (canSave, message) = (false, CheckBalance);
        else if (expired) (canSave, message) = (false, Expired);
        else if (errors.Count == 0) message = ClinicInvoiceService.Verified;
        return new MemberVerifyResult(errors, duplicate, duplicate is not null && line.Total < 0 && !r.ConfirmReversal, message, canSave);
    }

    private static string First(string? v) => string.IsNullOrWhiteSpace(v) ? "" : v.Trim()[..1].ToUpperInvariant();
}
