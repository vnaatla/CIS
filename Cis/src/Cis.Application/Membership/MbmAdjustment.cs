using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Membership;

/// <summary>A covered person of the member (showMBMCov, lstCovAdd).</summary>
/// <param name="CoverId">Cover id.</param>
/// <param name="Name">Name.</param>
/// <param name="Ic">IC.</param>
/// <param name="Relationship">REL code.</param>
/// <param name="Status">MBMCCLNStatus.</param>
/// <param name="Exclusion">MBMCCLNWarranty.</param>
/// <param name="Remarks">MBMCCLNRemarks.</param>
/// <param name="Dob">Date of birth.</param>
/// <param name="Plan">Cover plan (the principal's when empty).</param>
/// <param name="CardType">Card type.</param>
/// <param name="Premium">Premium.</param>
/// <param name="PremReduction">Premium reduction (cancelled memberships, MBMCISRefund).</param>
/// <param name="PremEffective">Effective premium.</param>
/// <param name="Mco">MCO fee.</param>
/// <param name="McoRefund">MCO refund.</param>
/// <param name="McoEffective">Effective MCO.</param>
/// <param name="AnnualLimit">Annual limit ("FULL" or amount).</param>
/// <param name="AccumAmt">Accumulated amount.</param>
/// <param name="BalanceLimit">Balance limit.</param>
/// <param name="VisitLimit">Visit limit.</param>
/// <param name="MepLimit">MEP limit.</param>
/// <param name="DentalLimit">Dental limit.</param>
/// <param name="DentalExtLimit">Dental extraction limit.</param>
/// <param name="SpecLimit">Spectacles limit.</param>
/// <param name="FirstSpecLimit">First spectacles limit.</param>
/// <param name="AgeCode">Age band code.</param>
/// <param name="AdultChild">A / C.</param>
public record AdjCover(string CoverId, string Name, string Ic, string Relationship, string Status, string Exclusion, string Remarks, DateTime? Dob,
    string Plan, string CardType, decimal Premium, decimal? PremReduction, decimal? PremEffective, decimal Mco, decimal? McoRefund, decimal? McoEffective,
    string AnnualLimit, decimal AccumAmt, string BalanceLimit, string VisitLimit, decimal MepLimit, decimal DentalLimit, decimal DentalExtLimit,
    decimal SpecLimit, decimal FirstSpecLimit, string AgeCode, string AdultChild);

/// <summary>The member shown by MBM Adjustment (showMBM).</summary>
/// <param name="MbmNumber">Member number.</param>
/// <param name="Status">MBMCLNStatus (A / C / D / S).</param>
/// <param name="ExpiryStatus">MBMCLNExpiryStatus.</param>
/// <param name="DataStatus">MBMCLNDataStatus.</param>
/// <param name="PayCode">Payor.</param>
/// <param name="Name">Name.</param>
/// <param name="PolicyNo">Policy no.</param>
/// <param name="Ic">IC.</param>
/// <param name="InsType">Insured type.</param>
/// <param name="Plan">Plan.</param>
/// <param name="Dob">Date of birth.</param>
/// <param name="AdultChild">A / C.</param>
/// <param name="GroupCompany">Group company.</param>
/// <param name="Department">Department.</param>
/// <param name="LimitStatus">Supplementary limit status code.</param>
/// <param name="LimitStatusText">Its text ("Share Limit For Family" …).</param>
/// <param name="EffDate">Policy eff date.</param>
/// <param name="ExpDate">Policy exp date.</param>
/// <param name="AgeCode">Age band code.</param>
/// <param name="CardType">Card type.</param>
/// <param name="Exclusion">MBMCLNWarranty.</param>
/// <param name="Remarks">MBMRemarks.</param>
/// <param name="CisRemarks">MBMCLNRemarks.</param>
/// <param name="Premium">Premium.</param>
/// <param name="McoFee">MCO fee.</param>
/// <param name="AccumAmt">Accumulated amount.</param>
/// <param name="Email">E-mail.</param>
/// <param name="Mobile">Mobile.</param>
/// <param name="BankName">Bank name.</param>
/// <param name="BankAccount">Bank account.</param>
/// <param name="AnnualLimit">Annual limit.</param>
/// <param name="VisitLimit">Visit limit.</param>
/// <param name="BalanceLimit">Balance limit.</param>
/// <param name="MepLimit">MEP limit.</param>
/// <param name="DentalLimit">Dental limit.</param>
/// <param name="SpecLimit">Spectacles limit.</param>
/// <param name="DentalExtLimit">Dental extraction limit.</param>
/// <param name="FirstSpecLimit">First spectacles limit.</param>
/// <param name="Covers">Covered persons (F / H only).</param>
/// <param name="CanEdit">Edit is allowed (not deleted / suspended).</param>
/// <param name="Notice">The desktop's message after loading a deleted / cancelled / suspended membership.</param>
public record AdjMember(string MbmNumber, string Status, string ExpiryStatus, string DataStatus, string PayCode, string Name, string PolicyNo, string Ic,
    string InsType, string Plan, DateTime? Dob, string AdultChild, string GroupCompany, string Department, string LimitStatus, string LimitStatusText,
    DateTime? EffDate, DateTime? ExpDate, string AgeCode, string CardType, string Exclusion, string Remarks, string CisRemarks, decimal Premium,
    decimal McoFee, decimal AccumAmt, string Email, string Mobile, string BankName, string BankAccount, string AnnualLimit, string VisitLimit,
    string BalanceLimit, decimal MepLimit, decimal DentalLimit, decimal SpecLimit, decimal DentalExtLimit, decimal FirstSpecLimit, List<AdjCover> Covers,
    bool CanEdit, string? Notice);

/// <summary>Combos.</summary>
/// <param name="CardTypes">CheckItems CSMD.</param>
/// <param name="Relationships">HISDB REL.</param>
public record AdjSetup(List<string> CardTypes, List<LookupOption> Relationships);

/// <summary>A premium / MCO look-up (GetPremium, GetMCO).</summary>
/// <param name="Amount">Amount.</param>
/// <param name="Found">False when the table had no row (the desktop's warning).</param>
public record FeeLookup(decimal Amount, bool Found);

/// <summary>Annual / visit limits (GetAnnualLimit, GetVisitLimit).</summary>
/// <param name="AnnualLimit">Annual limit.</param>
/// <param name="AnnualFull">"FULL".</param>
/// <param name="AnnualOk">False when AnnualLimit.SuppLimitStatus is empty (the desktop's warning).</param>
/// <param name="VisitLimit">Visit limit.</param>
/// <param name="VisitFull">"FULL".</param>
/// <param name="Status">The AnnualLimit status it read.</param>
public record LimitLookup(decimal AnnualLimit, bool AnnualFull, bool AnnualOk, decimal VisitLimit, bool VisitFull, string Status);

/// <summary>Fee recalculation for a new policy period (EffdateExpDateChg).</summary>
/// <param name="Member">Member number.</param>
/// <param name="EffDate">New eff date.</param>
/// <param name="ExpDate">New exp date.</param>
public record AdjFeesRequest(string? Member, string? EffDate, string? ExpDate);

/// <summary>New premium / MCO of the principal or a cover.</summary>
/// <param name="CoverId">"00" for the principal.</param>
/// <param name="AgeCode">Age band code.</param>
/// <param name="AdultChild">A / C.</param>
/// <param name="Premium">Premium.</param>
/// <param name="Mco">MCO for the period.</param>
public record AdjFee(string CoverId, string AgeCode, string AdultChild, decimal Premium, decimal Mco);

/// <summary>Edits of one cover (CmdUpdate / Cancel Supp).</summary>
/// <param name="CoverId">Cover id.</param>
/// <param name="Exclusion">Exclusion.</param>
/// <param name="Remarks">Remarks.</param>
/// <param name="Plan">Plan.</param>
/// <param name="CardType">Card type.</param>
/// <param name="Premium">Premium.</param>
/// <param name="Mco">MCO.</param>
/// <param name="Cancel">Flagged by Cancel Supp (K).</param>
public record AdjCoverEdit(string CoverId, string? Exclusion, string? Remarks, string? Plan, string? CardType, decimal Premium, decimal Mco, bool Cancel);

/// <summary>CmdSave_Click.</summary>
/// <param name="Member">Member number.</param>
/// <param name="EndtType">A / C / K / X / Y / D.</param>
/// <param name="EndtBordxDate">Endt. Bordx Date.</param>
/// <param name="EndtEffDate">Endt. Effective Date.</param>
/// <param name="PolicyNo">Policy no (A).</param>
/// <param name="EffDate">Policy eff date (A).</param>
/// <param name="ExpDate">Policy exp date (A).</param>
/// <param name="CisRemarks">CIS remarks (A, saved as MBMRemarks).</param>
/// <param name="CardType">Card type (A).</param>
/// <param name="McoFee">MCO fee (A).</param>
/// <param name="Premium">Premium (A).</param>
/// <param name="DataStatus">Data status (A).</param>
/// <param name="Status">Status (A).</param>
/// <param name="Exclusion">Exclusion (A).</param>
/// <param name="Plan">Plan (A).</param>
/// <param name="Email">E-mail (A).</param>
/// <param name="Mobile">Mobile (A).</param>
/// <param name="BankName">Bank name (A).</param>
/// <param name="BankAccount">Bank account (A).</param>
/// <param name="Covers">Cover edits (A, K).</param>
public record AdjSaveRequest(string? Member, string? EndtType, string? EndtBordxDate, string? EndtEffDate, string? PolicyNo, string? EffDate, string? ExpDate,
    string? CisRemarks, string? CardType, decimal McoFee, decimal Premium, string? DataStatus, string? Status, string? Exclusion, string? Plan, string? Email,
    string? Mobile, string? BankName, string? BankAccount, List<AdjCoverEdit>? Covers);

/// <summary>Lengths.</summary>
public sealed class AdjSaveRequestValidator : RequestValidator<AdjSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public AdjSaveRequestValidator()
    {
        RuleFor(r => r.EndtType).Must(t => t is null || t.Trim() is "A" or "C" or "K" or "X" or "Y" or "D").WithMessage("Unknown endorsement status.");
        RuleFor(r => r.PolicyNo).MaximumLength(100);
        RuleFor(r => r.Email).MaximumLength(200);
        RuleFor(r => r.Mobile).MaximumLength(50);
    }
}

/// <summary>Fee request validator.</summary>
public sealed class AdjFeesRequestValidator : RequestValidator<AdjFeesRequest>
{
    /// <summary>Defines the rules.</summary>
    public AdjFeesRequestValidator() => RuleFor(r => r.Member).MaximumLength(30);
}

/// <summary>Data access of MBM Adjustment (<c>cis.MbmAdj_*</c> and the shared member procedures).</summary>
public interface IMbmAdjustmentRepository
{
    /// <summary>Combos.</summary>
    Task<AdjSetup> SetupAsync(IDbSession session);

    /// <summary>SearchCISMemberProc rows.</summary>
    Task<List<Dictionary<string, object?>>> MemberAsync(IDbSession session, string mbmNumber);

    /// <summary>MBMOthers group company / department.</summary>
    Task<(string GroupCompany, string Department)> OthersAsync(IDbSession session, string mbmNumber);

    /// <summary>Supplementary limit status of plan + insured type (none = N).</summary>
    Task<string> LimitStatusAsync(IDbSession session, string plan, string insType);

    /// <summary>The cover's limit status on the policy date (none = A).</summary>
    Task<string> CoverLimitStatusAsync(IDbSession session, string plan, string insType, DateTime effDate);

    /// <summary>GetPremium.</summary>
    Task<FeeLookup> PremiumAsync(IDbSession session, string mbmType, string insType, string payCode, string ageCode, string plan, DateTime effDate);

    /// <summary>GetMCO.</summary>
    Task<FeeLookup> McoAsync(IDbSession session, string mbmType, string insType, string adultChild, DateTime effDate, string plan, string payCode);

    /// <summary>GetAnnualLimit + GetVisitLimit.</summary>
    Task<LimitLookup> LimitsAsync(IDbSession session, string mbmType, string insType, string plan, DateTime effDate, string payCode, string limitStatus);

    /// <summary>GetAgeBand code ("00" when none).</summary>
    Task<string> AgeCodeAsync(IDbSession session, int age, bool days);

    /// <summary>MBMCISRefund rows by cover.</summary>
    Task<Dictionary<string, (decimal PremReduction, decimal PremEffective, decimal McoRefund, decimal McoEffective)>> RefundsAsync(IDbSession session, string mbmNumber);

    /// <summary>Next / previous clinic member number.</summary>
    Task<string?> StepAsync(IDbSession session, string mbmNumber, bool next);

    /// <summary>Adjustment history tab.</summary>
    Task<List<Dictionary<string, object?>>> AmendHistoryAsync(IDbSession session, string mbmNumber);

    /// <summary>MBM history tab (SearchCISMBMHistory).</summary>
    Task<List<Dictionary<string, object?>>> MbmHistoryAsync(IDbSession session, string ic);

    /// <summary>An outstanding case exists.</summary>
    Task<bool> OpenCaseAsync(IDbSession session, string mbmNumber);

    /// <summary>Write2MBMnHistory.</summary>
    Task EndorseAsync(IDbSession session, string mbmNumber, string endtType, string status, DateTime bordxDate, DateTime endtEffDate, string userCode, DateTime today);

    /// <summary>WriteAdjStatus2Cov.</summary>
    Task CoverStatusAsync(IDbSession session, string mbmNumber, string status, string? covId, DateTime today);

    /// <summary>MemberMCORefund's MBMCISRefund row.</summary>
    Task RefundAsync(IDbSession session, string mbmNumber, string covId, decimal premEffective, decimal premReduction, decimal mcoEffective, decimal mcoRefund);

    /// <summary>SaveCLNOthers + SaveMBMTwo + SaveMBMOthersTwo.</summary>
    Task SavePrincipalAsync(IDbSession session, string mbmNumber, AdjSaveRequest r, DateTime? effDate, DateTime? expDate);

    /// <summary>SaveCisCovPersons (one cover).</summary>
    Task SaveCoverAsync(IDbSession session, string mbmNumber, AdjCoverEdit c, DateTime? effDate, DateTime? expDate);
}

/// <summary>Membership > MBM Adjustment.</summary>
public interface IMbmAdjustmentService
{
    /// <summary>ComboListing.</summary>
    Task<AdjSetup> SetupAsync();

    /// <summary>cmdShow_Click → FindMBM.</summary>
    Task<AdjMember> FindAsync(string? member);

    /// <summary>CmdPrev_Click / CmdNext_Click.</summary>
    Task<string> StepAsync(string? member, bool next);

    /// <summary>Adjustment history tab.</summary>
    Task<List<Dictionary<string, object?>>> AmendHistoryAsync(string? member);

    /// <summary>MBM history tab.</summary>
    Task<List<Dictionary<string, object?>>> MbmHistoryAsync(string? ic);

    /// <summary>EffdateExpDateChg: new premium / MCO for the principal and each cover, plus the desktop's look-up warnings.</summary>
    Task<(List<AdjFee> Fees, List<string> Warnings)> FeesAsync(AdjFeesRequest request);

    /// <summary>CmdSave_Click.</summary>
    Task<string> SaveAsync(AdjSaveRequest request, string userCode);
}

/// <inheritdoc />
public sealed class MbmAdjustmentService(IDbSessions sessions, IMbmAdjustmentRepository repository, TimeProvider clock) : IMbmAdjustmentService
{
    /// <summary>Desktop texts (spacing kept).</summary>
    public const string EnterMember = "Please Enter Membership No", NotFound = "No record found!", SelectFirst = "Please select a record to edit!",
        BordxFirst = "Please Enter Endt. Bordx Date!", EffFirst = "Please Enter Endt. Effective Date!", OpenCase = "There is at least an outstanding case under this member!",
        Deleted = "Membership has been deleted. No further changes are allowed!", Cancelled = "Membership has been canceled. No further changes are allowed!",
        Suspended = "Membership has been Suspended. No further changes are allowed!", NoPremium = "Age Out of Premium Range! Please check your Premium Table.",
        NoMco = "MCO Information Is Not Correct\n Please Check your MCO table", NoAnnual = "Annual Limit Information Is Not Correct! \nPlease check you Maintenance Table",
        Last = " Last Record Reached! ", First = "First Record Reached! \nNot a valid Membership Number Format! ";

    /// <summary>The message CmdSave_Click shows per endorsement type.</summary>
    public static string SavedMessage(string endtType) => endtType switch
    {
        "X" => "Record is Suspended",
        "D" => "Policy Deleted Successfully",
        "Y" => "Record is Activated",
        "C" => "This Membership has been cancelled successfully",
        "K" => "Supplementary has been cancelled successfully",
        _ => "Updated Successfully",
    };

    /// <summary>showMBM's limit status texts.</summary>
    public static string LimitStatusText(string code) => code switch
    {
        "N" => "Unlimited Limit",
        "A" => "Share Limit For Family",
        "B" => "Share Limit For Supplementary",
        "C" => "Separate Limit For Supplementary",
        "D" => "Share For Supplementary and Full For Principle",
        "E" => "Separate For Supplementary and Full For Principle",
        _ => code,
    };

    /// <inheritdoc />
    public async Task<AdjSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<AdjMember> FindAsync(string? member)
    {
        var number = Number(member);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.MemberAsync(s, number);
        if (rows.Count == 0) throw new NotFoundException(NotFound);
        var p = rows[0];
        var (group, dept) = await repository.OthersAsync(s, number);
        string ins = T(p, "INSType"), plan = T(p, "PLNType"), pay = T(p, "PAYCode"), status = T(p, "MBMCLNStatus");
        var eff = p.GetValueOrDefault("MBMPolEffDate") as DateTime?;
        var dob = p.GetValueOrDefault("MBMDOB") as DateTime?;
        var limitStatus = await repository.LimitStatusAsync(s, plan, ins);
        var limits = eff is null ? null : await repository.LimitsAsync(s, "P", ins, plan, eff.Value, pay, limitStatus);
        string ann = Limit(limits?.AnnualFull, limits?.AnnualLimit), visit = Limit(limits?.VisitFull, limits?.VisitLimit);
        var covers = new List<AdjCover>();
        if (ins is "F" or "H")
        {
            var refunds = status == "C" ? await repository.RefundsAsync(s, number) : [];
            foreach (var c in rows.Where(r => T(r, "MBMCCoverID") is { Length: > 0 } id && id != "00"))
                covers.Add(await Cover(s, p, c, ins, pay, eff, refunds));
        }
        var notice = status switch { "D" => Deleted, "C" => Cancelled, "S" => Suspended, _ => null };
        return new AdjMember(number, status, T(p, "MBMCLNExpiryStatus"), T(p, "MBMCLNDataStatus"), pay, T(p, "MBMName"), T(p, "MBMPolicyNo"), T(p, "MBMIcBcPp"),
            ins, plan, dob, T(p, "MBMAdultChild"), group, dept, limitStatus, LimitStatusText(limitStatus), eff, p.GetValueOrDefault("MBMPolExpDate") as DateTime?,
            eff is null || dob is null ? "00" : await AgeCode(s, eff.Value, dob.Value), T(p, "MBMCardType"), T(p, "MBMCLNWarranty"), T(p, "MBMRemarks"),
            T(p, "MBMCLNRemarks"), N(p, "MBMCLNPremium"), N(p, "MBMCLNMCOFee"), N(p, "MBMAccumAmt"), T(p, "MBMEmail"), T(p, "MBMTelNoM"), T(p, "MBMBankName"),
            T(p, "MBMBankACNo"), ann, visit == "FULL" ? visit : Money(N(p, "MBMCLNVisitLmt")), ann == "FULL" ? ann : Money(N(p, "MBMCLNBalLmt")),
            N(p, "MBMMEPLmt"), N(p, "MBMDentalLmt"), N(p, "MBMSpecLmt"), N(p, "MBMDentalExtLmt"), N(p, "MBMFirstSpecLmt"), covers, status is not ("D" or "S"), notice);
    }

    /// <inheritdoc />
    public async Task<string> StepAsync(string? member, bool next)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.StepAsync(s, (member ?? "").Trim().ToUpperInvariant(), next) ?? throw new NotFoundException(next ? Last : First);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> AmendHistoryAsync(string? member)
    {
        var number = Number(member);
        await using var s = await sessions.OpenAsync();
        return await repository.AmendHistoryAsync(s, number);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> MbmHistoryAsync(string? ic)
    {
        if (string.IsNullOrWhiteSpace(ic)) return [];
        await using var s = await sessions.OpenAsync();
        return await repository.MbmHistoryAsync(s, ic.Trim());
    }

    /// <inheritdoc />
    public async Task<(List<AdjFee> Fees, List<string> Warnings)> FeesAsync(AdjFeesRequest r)
    {
        var number = Number(r.Member);
        var eff = Date(r.EffDate) ?? throw new BusinessRuleException("Invalid Date Format! ");
        var exp = Date(r.ExpDate) ?? throw new BusinessRuleException("Invalid Date Format! ");
        await using var s = await sessions.OpenAsync();
        var rows = await repository.MemberAsync(s, number);
        if (rows.Count == 0) throw new NotFoundException(NotFound);
        var p = rows[0];
        string ins = T(p, "INSType"), plan = T(p, "PLNType"), pay = T(p, "PAYCode");
        var warnings = new List<string>();
        var fees = new List<AdjFee>();
        async Task Add(string covId, string type, DateTime? dob)
        {
            var age = dob is null ? "00" : await AgeCode(s, eff, dob.Value);
            var ac = age == "01" ? "C" : "A";
            var prem = await repository.PremiumAsync(s, type, ins, pay, age, plan, eff);
            var mco = await repository.McoAsync(s, type, ins, ac, eff, plan, pay);
            if (!prem.Found) warnings.Add(NoPremium);
            if (!mco.Found) warnings.Add(NoMco);
            fees.Add(new AdjFee(covId, age, ac, R2(prem.Amount), R2(InsertMcoFees(exp, eff, R2(mco.Amount)))));
        }
        // MemberCallFees uses the principal's plan for the covers too (cboPLNCode).
        await Add("00", "P", p.GetValueOrDefault("MBMDOB") as DateTime?);
        if (ins is "F" or "H")
            foreach (var c in rows.Where(x => T(x, "MBMCCoverID") is { Length: > 0 } id && id != "00"))
                await Add(T(c, "MBMCCoverID"), "S", c.GetValueOrDefault("MBMCDOB") as DateTime?);
        return (fees, warnings.Distinct().ToList());
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(AdjSaveRequest r, string userCode)
    {
        if (string.IsNullOrWhiteSpace(r.Member)) throw new BusinessRuleException(SelectFirst);
        var bordx = Date(r.EndtBordxDate) ?? throw new BusinessRuleException(BordxFirst);
        var endtEff = Date(r.EndtEffDate) ?? throw new BusinessRuleException(EffFirst);
        var type = (r.EndtType ?? "A").Trim();
        var number = Number(r.Member);
        var today = clock.GetLocalNow().Date;
        await using var s = await sessions.BeginAsync();
        var rows = await repository.MemberAsync(s, number);
        if (rows.Count == 0) throw new NotFoundException(NotFound);
        var p = rows[0];
        var status = T(p, "MBMCLNStatus");
        if (status == "D") throw new BusinessRuleException(Deleted);
        if (status == "S") throw new BusinessRuleException(Suspended);
        switch (type)
        {
            case "C":
                if (await repository.OpenCaseAsync(s, number)) throw new BusinessRuleException(OpenCase);
                await repository.EndorseAsync(s, number, type, "C", bordx, endtEff, userCode, today);
                await repository.CoverStatusAsync(s, number, "C", null, today);
                await Refund(s, p, "P", "00", T(p, "PLNType"), p.GetValueOrDefault("MBMDOB") as DateTime?, endtEff);
                foreach (var c in Covers(rows, p))
                    await Refund(s, p, "S", T(c, "MBMCCoverID"), CoverPlan(p, c), c.GetValueOrDefault("MBMCDOB") as DateTime?, endtEff);
                break;
            case "K":
                foreach (var c in Covers(rows, p).Where(c => r.Covers?.Any(e => e.Cancel && e.CoverId == T(c, "MBMCCoverID")) == true))
                {
                    await Refund(s, p, "S", T(c, "MBMCCoverID"), CoverPlan(p, c), c.GetValueOrDefault("MBMCDOB") as DateTime?, endtEff);
                    await repository.CoverStatusAsync(s, number, "C", T(c, "MBMCCoverID"), today);
                }
                break;
            case "X":
                await repository.EndorseAsync(s, number, type, "S", bordx, endtEff, userCode, today);
                await repository.CoverStatusAsync(s, number, "S", null, today);
                break;
            case "Y":
                await repository.EndorseAsync(s, number, type, "A", bordx, endtEff, userCode, today);
                await repository.CoverStatusAsync(s, number, "A", null, today);
                break;
            case "D":
                await repository.EndorseAsync(s, number, type, "D", bordx, endtEff, userCode, today);
                await repository.CoverStatusAsync(s, number, "D", null, today);
                break;
            default:
                // Amendment: Write2MBMnHistory("A") (status A, MBMCancelDate = endt. eff date, as the desktop), then the edits.
                await repository.EndorseAsync(s, number, type, "A", bordx, endtEff, userCode, today);
                var (eff, exp) = (Date(r.EffDate), Date(r.ExpDate));
                await repository.SavePrincipalAsync(s, number, r, eff, exp);
                foreach (var c in r.Covers ?? []) await repository.SaveCoverAsync(s, number, c, eff, exp);
                break;
        }
        await s.CommitAsync();
        return SavedMessage(type);
    }

    private async Task<AdjCover> Cover(IDbSession s, Dictionary<string, object?> p, Dictionary<string, object?> c, string ins, string pay, DateTime? eff,
        Dictionary<string, (decimal PremReduction, decimal PremEffective, decimal McoRefund, decimal McoEffective)> refunds)
    {
        var plan = CoverPlan(p, c);
        var covId = T(c, "MBMCCoverID");
        var suppStatus = eff is null ? "A" : await repository.CoverLimitStatusAsync(s, plan, ins, eff.Value);
        var limits = eff is null ? null : await repository.LimitsAsync(s, "S", ins, T(p, "PLNType"), eff.Value, pay, suppStatus);
        string ann = Limit(limits?.AnnualFull, limits?.AnnualLimit), visit = Limit(limits?.VisitFull, limits?.VisitLimit);
        var shared = suppStatus == "A";
        var dob = c.GetValueOrDefault("MBMCDOB") as DateTime?;
        var age = eff is null || dob is null ? "00" : await AgeCode(s, eff.Value, dob.Value);
        refunds.TryGetValue(covId, out var rf);
        var hasRefund = refunds.ContainsKey(covId);
        return new AdjCover(covId, T(c, "MBMCName"), T(c, "MBMCICBCPPNo"), T(c, "MBMCRELCode"), T(c, "MBMCCLNStatus"), T(c, "MBMCCLNWarranty"),
            T(c, "MBMCCLNRemarks"), dob, plan, T(c, "MBMCCardType"), N(c, "MBMCCLNPrem"), hasRefund ? rf.PremReduction : null, hasRefund ? rf.PremEffective : null,
            N(c, "MBMCCLNMCO"), hasRefund ? rf.McoRefund : null, hasRefund ? rf.McoEffective : null, ann, shared ? N(p, "MBMAccumAmt") : N(c, "MBMCAccumAmt"),
            ann == "FULL" ? ann : Money(shared ? N(p, "MBMCLNBalLmt") : N(c, "MBMCCLNBalLmt")),
            visit == "FULL" ? visit : Money(shared ? N(p, "MBMCLNVisitLmt") : N(c, "MBMCCLNVisitLmt")), N(c, "MBMCCLNMEPLmt"), N(c, "MBMCCLNDentalLmt"),
            N(c, "MBMCCLNDentalExtLmt"), N(c, "MBMCCLNSpecLmt"), N(c, "MBMCCLNFirstSpecLmt"), age, age == "01" ? "C" : "A");
    }

    /// <summary>MemberMCORefund: prorated MCO / premium for the remaining period after the cancellation date, written to MBMCISRefund.</summary>
    private async Task Refund(IDbSession s, Dictionary<string, object?> p, string type, string covId, string plan, DateTime? dob, DateTime cancelDate)
    {
        if (p.GetValueOrDefault("MBMPolEffDate") is not DateTime eff || p.GetValueOrDefault("MBMPolExpDate") is not DateTime exp) return;
        string ins = T(p, "INSType"), pay = T(p, "PAYCode");
        var age = dob is null ? "00" : await AgeCode(s, eff, dob.Value);
        var ac = age == "01" ? "C" : "A";
        var r = Calculate(eff, exp, cancelDate, (await repository.McoAsync(s, type, ins, ac, eff, plan, pay)).Amount,
            (await repository.PremiumAsync(s, type, ins, pay, age, plan, eff)).Amount);
        await repository.RefundAsync(s, T(p, "MBMNumber"), covId, r.PremEffective, r.PremReduction, r.McoEffective, r.McoRefund);
    }

    /// <summary>Medix.bas InsertMCOFees / MCORefund / InsertPremium / MCOPremiumReduction as MemberMCORefund chains them (2-decimal rounding between steps).</summary>
    public static (decimal PremEffective, decimal PremReduction, decimal McoEffective, decimal McoRefund) Calculate(DateTime eff, DateTime exp, DateTime cancel,
        decimal basicMco, decimal basicPremium)
    {
        var mcoFee = InsertMcoFees(exp, eff, R2(basicMco));
        var mcoRed = R2(McoRefund(exp, cancel, eff, mcoFee));
        var premium = InsertPremium(exp, eff, R2(basicPremium));
        var premRed = R2(PremiumReduction(exp, cancel, eff, premium));
        return (R2(premium - premRed), premRed, R2(mcoFee - mcoRed), mcoRed);
    }

    /// <summary>InsertMCOFees: days of the period × fee / 364.</summary>
    public static decimal InsertMcoFees(DateTime exp, DateTime eff, decimal fee) => (decimal)(exp - eff).TotalDays * fee / 364m;

    /// <summary>InsertPremium: days / 364 × premium.</summary>
    public static decimal InsertPremium(DateTime exp, DateTime eff, decimal premium) => (decimal)(exp - eff).TotalDays / 364m * premium;

    /// <summary>MCORefund: remaining days / period days × fee, never below 0.</summary>
    public static decimal McoRefund(DateTime exp, DateTime cancel, DateTime eff, decimal fee)
    {
        var period = (decimal)(exp - eff).TotalDays;
        return period == 0 ? 0 : Math.Max(0, (decimal)(exp - cancel).TotalDays / period * fee);
    }

    /// <summary>MCOPremiumReduction: remaining days / period days × premium when cancelled before the expiry.</summary>
    public static decimal PremiumReduction(DateTime exp, DateTime cancel, DateTime eff, decimal premium)
    {
        var period = (decimal)(exp - eff).TotalDays;
        return cancel < exp && period != 0 ? (decimal)(exp - cancel).TotalDays / period * premium : 0;
    }

    private async Task<string> AgeCode(IDbSession s, DateTime eff, DateTime dob)
    {
        var age = MemberRegistrationService.GetAge(eff, dob);
        return await repository.AgeCodeAsync(s, int.Parse(age[..^1], CultureInfo.InvariantCulture), age.EndsWith('D'));
    }

    private static IEnumerable<Dictionary<string, object?>> Covers(List<Dictionary<string, object?>> rows, Dictionary<string, object?> p) =>
        T(p, "INSType") is "F" or "H" ? rows.Where(r => T(r, "MBMCCoverID") is { Length: > 0 } id && id != "00") : [];

    private static string CoverPlan(Dictionary<string, object?> p, Dictionary<string, object?> c) =>
        T(c, "MBMCCLNPLNCode") is { Length: > 0 } plan ? plan : T(p, "PLNType");

    private static string Number(string? member) =>
        string.IsNullOrWhiteSpace(member) ? throw new BusinessRuleException(EnterMember) : member.Trim().ToUpperInvariant();

    private static string Limit(bool? full, decimal? amount) => full == true ? "FULL" : Money(amount ?? 0);

    private static string Money(decimal v) => v.ToString("#,##0.00", CultureInfo.InvariantCulture);

    private static decimal R2(decimal v) => Math.Round(v, 2, MidpointRounding.AwayFromZero);

    private static DateTime? Date(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static string T(Dictionary<string, object?> row, string key) =>
        row.TryGetValue(key, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(Dictionary<string, object?> row, string key) => row.GetValueOrDefault(key) switch
    {
        decimal d => d,
        double d => (decimal)d,
        int i => i,
        string t when decimal.TryParse(t, NumberStyles.Number, CultureInfo.InvariantCulture, out var d) => d,
        _ => 0,
    };
}
