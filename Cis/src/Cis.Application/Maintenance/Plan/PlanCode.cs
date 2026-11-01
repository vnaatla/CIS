using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.Plan;

/// <summary>Plan Code tab fields (frmMaintenancePlan tab 2). Combos carry their text or code; the service cuts them like SavePLN.</summary>
/// <param name="HealthType">HLT Type (S / N …).</param>
/// <param name="PayCode">CIS Pay Code.</param>
/// <param name="HisPayCode">HIS Pay Code.</param>
/// <param name="GroupCompany">Grp Comp. ("CODE - NAME").</param>
/// <param name="ClnContract">Cln Cont. (CheckItems CSMD).</param>
/// <param name="AnnLmt">Principal Ann Lmt (Yes / No).</param>
/// <param name="VisLmt">Principal Vst Lmt.</param>
/// <param name="Prm">Prem.</param>
/// <param name="Mco">M C O.</param>
/// <param name="SAnnLmt">Supplementary Ann Lmt.</param>
/// <param name="SVisLmt">Supplementary Vst Lmt.</param>
/// <param name="DcStatus">DC Status (N / Y).</param>
/// <param name="NoOfPlans">No of Plan Code (1–6, health type N).</param>
/// <param name="PlanCodes">Plan Code boxes 1–6 (health type S; box 1 is also a search filter).</param>
/// <param name="Descriptions">Plan Description boxes 1–6.</param>
/// <param name="Insured">Insured (Y / N).</param>
/// <param name="HrApproval">HR approval (Y / N).</param>
/// <param name="OpFamilyStatus">OP Family Status (B-Both / D-Dependent Only / E-Employee Only), stored as shown.</param>
/// <param name="PrinSpLimit">Principal SP Limit.</param>
/// <param name="PrinGpLimit">Principal GP Limit.</param>
/// <param name="DepSpLimit">Dependent SP Limit.</param>
/// <param name="DepGpLimit">Dependent GP Limit.</param>
/// <param name="SpecialRemark">Special Remark.</param>
/// <param name="OpDependent">OP Dependent (N / Y).</param>
/// <param name="PDentalStatus">Principal Dental Status (search only; Save writes N).</param>
/// <param name="PDentalLimit">Principal Dental Limit (search only).</param>
/// <param name="PSpecStatus">Principal Spect Status (search only).</param>
/// <param name="PSpecLimit">Principal Spect Limit (search only).</param>
/// <param name="PMepStatus">Principal MEP Status (search only).</param>
/// <param name="SDentalStatus">Supplementary Dental Status (search only).</param>
/// <param name="SDentalLimit">Supplementary Dental Limit (search only).</param>
/// <param name="SSpecStatus">Supplementary Spec. Status (search only).</param>
/// <param name="SSpecLimit">Supplementary Spec. Limit (search only).</param>
/// <param name="SMepStatus">Supplementary MEP Status (search only).</param>
public record PlanCodeDto(string? HealthType, string? PayCode, string? HisPayCode, string? GroupCompany, string? ClnContract, string? AnnLmt,
    string? VisLmt, string? Prm, string? Mco, string? SAnnLmt, string? SVisLmt, string? DcStatus, string? NoOfPlans, List<string?>? PlanCodes,
    List<string?>? Descriptions, string? Insured, string? HrApproval, string? OpFamilyStatus, string? PrinSpLimit, string? PrinGpLimit,
    string? DepSpLimit, string? DepGpLimit, string? SpecialRemark, string? OpDependent, string? PDentalStatus = null, string? PDentalLimit = null,
    string? PSpecStatus = null, string? PSpecLimit = null, string? PMepStatus = null, string? SDentalStatus = null, string? SDentalLimit = null,
    string? SSpecStatus = null, string? SSpecLimit = null, string? SMepStatus = null);

/// <summary>Save / search of the Plan Code tab.</summary>
/// <param name="Plan">The fields.</param>
public record PlanCodeRequest(PlanCodeDto Plan);

/// <summary>The search filters as SearchPLN put them into SQL.</summary>
/// <param name="PlnCode">PLNCode LIKE.</param>
/// <param name="PayCode">PAYCode =.</param>
/// <param name="HisPayCode">HISPAYCode =.</param>
/// <param name="HltCode">HLTCode =.</param>
/// <param name="GrpCode">CoGRPCode =.</param>
/// <param name="ClnContract">ClnContract =.</param>
/// <param name="Flags">The one-letter indicator filters by procedure parameter name.</param>
public record PlanCodeFilters(string? PlnCode, string? PayCode, string? HisPayCode, string? HltCode, string? GrpCode, string? ClnContract,
    IReadOnlyDictionary<string, string?> Flags);

/// <summary>A PLN row as written by SavePLN (N).</summary>
/// <param name="Code">PLNCode.</param>
/// <param name="Desc">PLNDesc.</param>
/// <param name="HltCode">HLTCode.</param>
/// <param name="PayCode">PAYCode.</param>
/// <param name="HisPayCode">HISPAYCode.</param>
/// <param name="GrpCompany">GRPCompany.</param>
/// <param name="AnnLmt">ANNLmtInd.</param>
/// <param name="VisLmt">VISLmtInd.</param>
/// <param name="SAnnLmt">ANNLmtIndSupp.</param>
/// <param name="SVisLmt">VISLmtIndSupp.</param>
/// <param name="Prm">PRMInd.</param>
/// <param name="Mco">MCOInd.</param>
/// <param name="ClnContract">ClnContract.</param>
/// <param name="DcStatus">DCStatus.</param>
/// <param name="Insured">Insured.</param>
/// <param name="HrApproval">HRApproval.</param>
/// <param name="OpFamilyStatus">OPFamilyStatus.</param>
/// <param name="PrinSpLimit">PrinSPLimit.</param>
/// <param name="PrinGpLimit">PrinGPLimit.</param>
/// <param name="DepSpLimit">DepSPLimit.</param>
/// <param name="DepGpLimit">DepGPLimit.</param>
/// <param name="SpecialRemark">PLNSpecialRemark.</param>
/// <param name="OpDependent">isOPDependent.</param>
public record PlanRow(string Code, string Desc, string HltCode, string PayCode, string HisPayCode, string GrpCompany, string AnnLmt, string VisLmt,
    string SAnnLmt, string SVisLmt, string Prm, string Mco, string ClnContract, string DcStatus, string Insured, string HrApproval,
    string OpFamilyStatus, double PrinSpLimit, double PrinGpLimit, double DepSpLimit, double DepGpLimit, string SpecialRemark, string OpDependent);

/// <summary>Combos of the tab.</summary>
/// <param name="HealthTypes">HLT.</param>
/// <param name="Payors">CIS PAY.</param>
/// <param name="HisPayors">HISDB PAY.</param>
/// <param name="ClinicContracts">CheckItems CSMD.</param>
/// <param name="GroupCompanies">GRPCompany (Check2).</param>
public record PlanCodeSetup(List<LookupOption> HealthTypes, List<LookupOption> Payors, List<LookupOption> HisPayors, List<LookupOption> ClinicContracts,
    List<LookupOption> GroupCompanies);

/// <summary>Column lengths of PLN (the desktop's MaxLength).</summary>
public sealed class PlanCodeRequestValidator : RequestValidator<PlanCodeRequest>
{
    /// <summary>Defines the rules.</summary>
    public PlanCodeRequestValidator() => RuleFor(r => r.Plan).NotNull().ChildRules(p =>
    {
        p.RuleForEach(x => x.PlanCodes).MaximumLength(4);
        p.RuleForEach(x => x.Descriptions).MaximumLength(200);
        p.RuleFor(x => x.SpecialRemark).MaximumLength(1500);
        p.RuleFor(x => x.ClnContract).MaximumLength(50);
        p.RuleFor(x => x.OpFamilyStatus).MaximumLength(50);
        p.RuleFor(x => x.GroupCompany).MaximumLength(300);
        p.RuleFor(x => x.NoOfPlans).Must(n => string.IsNullOrWhiteSpace(n) || n.Trim() is "1" or "2" or "3" or "4" or "5" or "6")
            .WithMessage("No of Plan Code must be 1 to 6.");
    });
}

/// <summary>Data access of the Plan Code tab.</summary>
public interface IPlanCodeRepository
{
    /// <summary>Combos.</summary>
    Task<PlanCodeSetup> SetupAsync(IDbSession session);

    /// <summary>CIS payors of a health type.</summary>
    Task<List<LookupOption>> PayorsAsync(IDbSession session, string hltCode);

    /// <summary>VIEW_PLNListing rows.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PlanCodeFilters filters);

    /// <summary>One plan.</summary>
    Task<Dictionary<string, object?>?> GetAsync(IDbSession session, string code);

    /// <summary>True when the text is a group company name.</summary>
    Task<bool> GroupNameExistsAsync(IDbSession session, string name);

    /// <summary>GeneratePLNSeq.</summary>
    Task<string> NextCodeAsync(IDbSession session, string letter);

    /// <summary>Insert one plan.</summary>
    Task InsertAsync(IDbSession session, PlanRow row, string userCode, DateTime today);

    /// <summary>True when membership uses the plan.</summary>
    Task<bool> InUseAsync(IDbSession session, string code);

    /// <summary>Delete; rows removed.</summary>
    Task<int> DeleteAsync(IDbSession session, string code);
}

/// <summary>Maintenance > Plan, Plan Code tab.</summary>
public interface IPlanCodeService
{
    /// <summary>PLNComboListing.</summary>
    Task<PlanCodeSetup> SetupAsync();

    /// <summary>cboPLNHLTCode_Click.</summary>
    Task<List<LookupOption>> PayorsAsync(string? hltCode);

    /// <summary>SearchPLN; "Record not found" when empty.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(PlanCodeDto filters);

    /// <summary>lstPlanList_Click.</summary>
    Task<Dictionary<string, object?>> GetAsync(string code);

    /// <summary>SavePLN; returns the message and the codes created.</summary>
    Task<(string Message, List<string> Codes)> SaveAsync(PlanCodeDto plan, string userCode);

    /// <summary>DeletePLN.</summary>
    Task<string> DeleteAsync(string? code);
}

/// <inheritdoc />
public sealed class PlanCodeService(IDbSessions sessions, IPlanCodeRepository repository, TimeProvider clock) : IPlanCodeService
{
    /// <summary>SavePLN's checks in order (texts kept, trailing spaces included); N / S are the first letter of the health type.</summary>
    public static readonly IReadOnlyList<(Func<PlanCodeDto, bool> Missing, string Message)> SaveChecks =
    [
        (p => E(p.HealthType), "Please Enter Health Code "),
        (p => Hlt(p) == "N" && E(p.PayCode), "Please Enter Payor Code "),
        (p => Hlt(p) == "N" && E(p.NoOfPlans), "Please Enter No of Plan Code "),
        (p => E(p.ClnContract), "Please Enter Clinic Contract"),
        (p => E(p.AnnLmt), "Please Enter Principal Annual Limit"),
        (p => E(p.VisLmt), "Please Enter Principal Visit Limit"),
        (p => E(p.Prm), "Please Enter Premium"),
        (p => E(p.Mco), "Please Enter MCO"),
        (p => E(p.SAnnLmt), "Please Enter Supplementary Annual Limit"),
        (p => E(p.SVisLmt), "Please Enter Supplementary Visit Limit"),
        (p => E(p.DcStatus), "Please Enter DC Status"),
        (p => E(p.GroupCompany), "Please Enter Group Company"),
        (p => Hlt(p) == "S" && E(p.PlanCodes?.ElementAtOrDefault(0)), "Please Enter Plan Code"),
        (p => E(p.OpDependent), "Please Select OP Dependent"),
    ];

    /// <summary>Texts.</summary>
    public const string GroupDuplicated = "Record Duplicated! Please change the corresponding field", NotFound = "Record not found",
        Saved = "Record Saved...", Deleted = "Record Deleted... ", SearchFirst = "Please search the record",
        InUse = "This record cannot be deleted because it is used in Membership table!",
        OnlyNonSihat = "Plans of this health type cannot be saved here: the desktop wrote them to the HIS plan table, which has none of these columns, so its save always failed.";

    /// <inheritdoc />
    public async Task<PlanCodeSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<LookupOption>> PayorsAsync(string? hltCode)
    {
        await using var s = await sessions.OpenAsync();
        return E(hltCode) ? [] : await repository.PayorsAsync(s, hltCode!.Trim()[..1]);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(PlanCodeDto filters)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, Reduce(filters));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>> GetAsync(string code)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.GetAsync(s, code.Trim()) ?? throw new NotFoundException(NotFound);
    }

    /// <inheritdoc />
    public async Task<(string Message, List<string> Codes)> SaveAsync(PlanCodeDto plan, string userCode)
    {
        foreach (var (missing, message) in SaveChecks)
            if (missing(plan)) throw new BusinessRuleException(message);
        if (Hlt(plan) != "N") throw new BusinessRuleException(OnlyNonSihat);
        await using var s = await sessions.BeginAsync();
        if (await repository.GroupNameExistsAsync(s, plan.GroupCompany!.Trim())) throw new ConflictException(GroupDuplicated);
        var count = int.Parse(plan.NoOfPlans!.Trim(), CultureInfo.InvariantCulture);
        var codes = new List<string>();
        var today = clock.GetLocalNow().Date;
        for (var i = 0; i < count; i++)
        {
            var code = await repository.NextCodeAsync(s, plan.GroupCompany.Trim()[..1].ToUpperInvariant());
            await repository.InsertAsync(s, Row(plan, code, plan.Descriptions?.ElementAtOrDefault(i)), userCode, today);
            codes.Add(code);
        }
        await s.CommitAsync();
        return (Saved, codes);
    }

    /// <inheritdoc />
    public async Task<string> DeleteAsync(string? code)
    {
        if (E(code)) throw new BusinessRuleException(SearchFirst);
        await using var s = await sessions.BeginAsync();
        if (await repository.InUseAsync(s, code!.Trim())) throw new ConflictException(InUse);
        if (await repository.DeleteAsync(s, code.Trim().ToUpperInvariant()) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return Deleted;
    }

    /// <summary>SavePLN (N) column values for one plan.</summary>
    public static PlanRow Row(PlanCodeDto p, string code, string? description) => new(
        code, U(description), Before(p.HealthType, 1), Before(p.PayCode, 4), Before(p.HisPayCode, 4), Before(p.GroupCompany, 4), F(p.AnnLmt), F(p.VisLmt),
        F(p.SAnnLmt), F(p.SVisLmt), F(p.Prm), F(p.Mco), (p.ClnContract ?? "").Trim(), F(p.DcStatus), F(p.Insured), F(p.HrApproval), p.OpFamilyStatus ?? "",
        Num(p.PrinSpLimit), Num(p.PrinGpLimit), Num(p.DepSpLimit), Num(p.DepGpLimit), U(p.SpecialRemark), F(p.OpDependent));

    /// <summary>SearchPLN's values: plan code upper-cased, pay codes Mid(…, 1, 2), health type / indicators first letter, group Mid(…, 1, 4).</summary>
    public static PlanCodeFilters Reduce(PlanCodeDto f)
    {
        static string? Cut(string? v, int n) => E(v) ? null : v!.Length > n ? v[..n] : v;
        static string? First(string? v) => E(v?.Trim()) ? null : v!.Trim()[..1].ToUpperInvariant();
        return new(E(f.PlanCodes?.ElementAtOrDefault(0)?.Trim()) ? null : f.PlanCodes![0]!.Trim().ToUpperInvariant(), Cut(f.PayCode?.Trim(), 2),
            Cut(f.HisPayCode?.Trim(), 2), First(f.HealthType), Cut(f.GroupCompany?.Trim(), 4), E(f.ClnContract?.Trim()) ? null : f.ClnContract!.Trim(),
            new Dictionary<string, string?>
            {
                ["annLmt"] = First(f.AnnLmt), ["visLmt"] = First(f.VisLmt), ["prm"] = First(f.Prm), ["mco"] = First(f.Mco),
                ["pDentalStatus"] = First(f.PDentalStatus), ["pDentalLimit"] = First(f.PDentalLimit), ["pSpecStatus"] = First(f.PSpecStatus),
                ["pSpecLimit"] = First(f.PSpecLimit), ["pMepStatus"] = First(f.PMepStatus), ["sDentalStatus"] = First(f.SDentalStatus),
                ["sDentalLimit"] = First(f.SDentalLimit), ["sSpecStatus"] = First(f.SSpecStatus), ["sSpecLimit"] = First(f.SSpecLimit),
                ["sMepStatus"] = First(f.SMepStatus), ["sAnnLmt"] = First(f.SAnnLmt), ["sVisLmt"] = First(f.SVisLmt),
            });
    }

    private static string? Hlt(PlanCodeDto p) => E(p.HealthType?.Trim()) ? null : p.HealthType!.Trim()[..1].ToUpperInvariant();

    private static bool E(string? v) => string.IsNullOrWhiteSpace(v);

    private static string U(string? v) => (v ?? "").Trim().ToUpperInvariant();

    private static string F(string? v) => E(v) ? "" : v!.Trim()[..1].ToUpperInvariant();

    private static double Num(string? v) => double.TryParse(v?.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0;

    /// <summary>Trim(Mid(text, 1, IIf(InStr(text, "-") = 0, n, InStr(text, "-") - 1))).</summary>
    private static string Before(string? text, int n)
    {
        if (E(text)) return "";
        var at = text!.IndexOf('-');
        return (at < 0 ? text[..Math.Min(n, text.Length)] : text[..at]).Trim();
    }
}
