using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.Plan;

/// <summary>Annual Limit tab fields (frmMaintenancePlan tab 4). Combos carry their code (or "CODE - NAME").</summary>
/// <param name="Index">AnnualIndex of the selected row (Edit / Delete).</param>
/// <param name="PlnCode">Plan.</param>
/// <param name="HltCode">Health Type.</param>
/// <param name="InsCode">Insured Type.</param>
/// <param name="PayCode">Payor (search only).</param>
/// <param name="GrpCompany">Group company (search only, prefix).</param>
/// <param name="AnnLimit">Principal annual limit.</param>
/// <param name="PDental">Principal dental annual limit.</param>
/// <param name="PSpec">Principal spectacles annual limit.</param>
/// <param name="PMep">Principal MEP annual limit.</param>
/// <param name="SuppStatus">A - Share Limit For Family · B - Share Limit For Supplementary · C - Separate Limit For Supplementary.</param>
/// <param name="SuppLimit">Supp premium limit (B / C).</param>
/// <param name="SDental">Supplementary dental annual limit (B / C).</param>
/// <param name="SSpec">Supplementary spectacles annual limit (B / C).</param>
/// <param name="SMep">Supplementary MEP annual limit (B / C).</param>
/// <param name="EffDate">Effective date dd/mm/yyyy.</param>
/// <param name="Version">Version.</param>
public record AnnualLimitDto(int? Index, string? PlnCode, string? HltCode, string? InsCode, string? PayCode, string? GrpCompany, string? AnnLimit,
    string? PDental, string? PSpec, string? PMep, string? SuppStatus, string? SuppLimit, string? SDental, string? SSpec, string? SMep, string? EffDate,
    string? Version);

/// <summary>Save / search of the Annual Limit tab.</summary>
/// <param name="Limit">The fields.</param>
public record AnnualLimitRequest(AnnualLimitDto Limit);

/// <summary>An AnnualLimit row as SaveANN writes it.</summary>
/// <param name="Index">AnnualIndex (null = Add).</param>
/// <param name="PlnCode">PLNCode.</param>
/// <param name="InsCode">INSCode.</param>
/// <param name="HltCode">HLTCode.</param>
/// <param name="AnnLimit">AnnLimit.</param>
/// <param name="PDental">PDentalAnnLimit.</param>
/// <param name="PSpec">PSpecAnnLimit.</param>
/// <param name="PMep">PMEPAnnLimit.</param>
/// <param name="Status">SuppLimitStatus.</param>
/// <param name="SuppLimit">SuppLimit.</param>
/// <param name="SDental">SDentalAnnLimit.</param>
/// <param name="SSpec">SSpecAnnLimit.</param>
/// <param name="SMep">SMEPAnnLimit.</param>
/// <param name="EffDate">AnnEffDate.</param>
/// <param name="Version">AnnLimitVersion.</param>
public record AnnualLimitRow(int? Index, string PlnCode, string InsCode, string HltCode, decimal AnnLimit, decimal PDental, decimal PSpec, decimal PMep,
    string Status, decimal? SuppLimit, decimal? SDental, decimal? SSpec, decimal? SMep, DateTime EffDate, int Version);

/// <summary>Combos of the Annual Limit tab.</summary>
/// <param name="Plans">PLN.</param>
/// <param name="InsuredTypes">INS.</param>
/// <param name="HealthTypes">HLT.</param>
/// <param name="Payors">PAY.</param>
/// <param name="GroupCompanies">Distinct PLN.GRPCompany.</param>
public record AnnualLimitSetup(List<LookupOption> Plans, List<LookupOption> InsuredTypes, List<LookupOption> HealthTypes, List<LookupOption> Payors,
    List<LookupOption> GroupCompanies);

/// <summary>Number formats of the limits.</summary>
public sealed class AnnualLimitRequestValidator : RequestValidator<AnnualLimitRequest>
{
    /// <summary>Defines the rules.</summary>
    public AnnualLimitRequestValidator() => RuleFor(r => r.Limit).NotNull().ChildRules(l =>
    {
        foreach (var (get, name) in new (System.Linq.Expressions.Expression<Func<AnnualLimitDto, string?>>, string)[]
                 {
                     (x => x.AnnLimit, "Principal Annual Limit"), (x => x.PDental, "Principal Dental Annual Limit"),
                     (x => x.PSpec, "Principal Spectacles Annual Limit"), (x => x.PMep, "Principal MEP Annual Limit"),
                     (x => x.SuppLimit, "Supp Premium Limit"), (x => x.SDental, "Supplementary Dental Annual Limit"),
                     (x => x.SSpec, "Supplementary Spectacles Annual Limit"), (x => x.SMep, "Supplementary MEP Annual Limit"),
                 })
            l.RuleFor(get).Must(v => string.IsNullOrWhiteSpace(v) || PlanAnnualLimitService.Money(v) is not null).WithMessage($"{name} must be a number.");
        l.RuleFor(x => x.Version).Must(v => string.IsNullOrWhiteSpace(v) || int.TryParse(v.Trim(), out _)).WithMessage("Annual Limit Version must be a whole number.");
    });
}

/// <summary>Data access of the Annual Limit tab.</summary>
public interface IPlanAnnualLimitRepository
{
    /// <summary>Combos.</summary>
    Task<AnnualLimitSetup> SetupAsync(IDbSession session);

    /// <summary>VIEW_ANN_PAY_PLN rows.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? pln, string? hlt, string? ins, string? pay, string? grp);

    /// <summary>True when insured + health + plan + version + date exists.</summary>
    Task<bool> ExistsAsync(IDbSession session, AnnualLimitRow row);

    /// <summary>Insert / update; rows written.</summary>
    Task<int> SaveAsync(IDbSession session, AnnualLimitRow row, string userCode, DateTime today);

    /// <summary>True when members use the plan (false where MBM does not exist).</summary>
    Task<bool> InUseAsync(IDbSession session, string plnCode);

    /// <summary>Delete; rows removed.</summary>
    Task<int> DeleteAsync(IDbSession session, int index);
}

/// <summary>Maintenance > Plan, Annual Limit tab.</summary>
public interface IPlanAnnualLimitService
{
    /// <summary>AnnualComboListing.</summary>
    Task<AnnualLimitSetup> SetupAsync();

    /// <summary>SearchANN; "Record not found" when empty.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(AnnualLimitDto filters);

    /// <summary>SaveANN.</summary>
    Task<string> SaveAsync(AnnualLimitDto limit, bool editing, string userCode);

    /// <summary>DeleteANN.</summary>
    Task<string> DeleteAsync(AnnualLimitDto limit);
}

/// <inheritdoc />
public sealed class PlanAnnualLimitService(IDbSessions sessions, IPlanAnnualLimitRepository repository, TimeProvider clock) : IPlanAnnualLimitService
{
    /// <summary>SaveANN's checks in order; B / C add the supplementary ones.</summary>
    public static readonly IReadOnlyList<(Func<AnnualLimitDto, bool> Missing, string Message)> SaveChecks =
    [
        (a => E(a.PlnCode), "Please Enter Plan Code"),
        (a => E(a.HltCode), "Please Enter Health Code"),
        (a => E(a.InsCode), "Please Enter Insured Code"),
        (a => E(a.AnnLimit), "Please Enter Principal Annual Limit"),
        (a => E(a.PDental), "Please Enter Principal Dental Annual Limit"),
        (a => E(a.PSpec), "Please Enter Principal Spectacles Annual Limit"),
        (a => E(a.PMep), "Please Enter Principal MEP Annual Limit"),
        (a => E(a.EffDate) || a.EffDate!.Trim() == "__/__/____", "Please Enter Annual Limit Effective Date"),
        (a => E(a.Version), "Please Enter Annual Limit Version"),
        (a => Supp(a) && E(a.SuppLimit), "Please Enter Supp Premium Limit"),
        (a => Supp(a) && E(a.SDental), "Please Enter Supplementary Dental Annual Limit"),
        (a => Supp(a) && E(a.SSpec), "Please Enter Supplementary Spectacles Annual Limit"),
        (a => Supp(a) && E(a.SMep), "Please Enter Supplementary MEP Annual Limit"),
    ];

    /// <summary>Desktop texts.</summary>
    public const string Duplicated = "Record Duplicated! \n Please Change The Corresponding Field", NotFound = "Record not found",
        Saved = "Record Saved...", Updated = "Record Updated...", Deleted = "Record Deleted...", SearchFirst = "Please search the record",
        PlanFirst = "Please Enter Insured Code", InUse = "This record cannot be deleted because it is used in Member Table!", BadDate = "Invalid Date Format! ";

    /// <inheritdoc />
    public async Task<AnnualLimitSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(AnnualLimitDto f)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, Opt(Before(f.PlnCode, 4)), Opt(Before(f.HltCode, 4)), Opt(Before(f.InsCode, 4)),
            Opt(Before(f.PayCode, 4)), Opt(Before(f.GrpCompany, 4)));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(AnnualLimitDto a, bool editing, string userCode)
    {
        foreach (var (missing, message) in SaveChecks)
            if (missing(a)) throw new BusinessRuleException(message);
        if (editing && a.Index is null) throw new BusinessRuleException(SearchFirst);
        var row = Row(a, editing);
        await using var s = await sessions.BeginAsync();
        if (!editing && await repository.ExistsAsync(s, row with { PlnCode = Head(a.PlnCode, 4) })) throw new ConflictException(Duplicated);
        if (await repository.SaveAsync(s, row, userCode, clock.GetLocalNow().Date) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return editing ? Updated : Saved;
    }

    /// <inheritdoc />
    public async Task<string> DeleteAsync(AnnualLimitDto a)
    {
        if (string.IsNullOrEmpty(a.PlnCode)) throw new BusinessRuleException(PlanFirst);
        if (a.Index is not { } index) throw new BusinessRuleException(SearchFirst);
        await using var s = await sessions.BeginAsync();
        if (await repository.InUseAsync(s, a.PlnCode.Trim())) throw new ConflictException(InUse);
        if (await repository.DeleteAsync(s, index) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return Deleted;
    }

    /// <summary>SaveANN's values: plan before "-" (else first 4), insured / health first character, the limits, status, supplementary
    /// limits only for B / C, date and version.</summary>
    public static AnnualLimitRow Row(AnnualLimitDto a, bool editing)
    {
        var status = Head(a.SuppStatus, 1) is { Length: 1 } st ? st : "A";
        var supp = status != "A";
        var eff = DateTime.TryParseExact(a.EffDate?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
            ? d : throw new BusinessRuleException(BadDate);
        return new(editing ? a.Index : null, Before(a.PlnCode, 4), Raw(a.InsCode, 1), Raw(a.HltCode, 1), Money(a.AnnLimit) ?? 0, Money(a.PDental) ?? 0,
            Money(a.PSpec) ?? 0, Money(a.PMep) ?? 0, status, supp ? Money(a.SuppLimit) : null, supp ? Money(a.SDental) : null,
            supp ? Money(a.SSpec) : null, supp ? Money(a.SMep) : null, eff, int.Parse(a.Version!.Trim(), CultureInfo.InvariantCulture));
    }

    /// <summary>A money box (thousands separators allowed, as the list shows them).</summary>
    public static decimal? Money(string? v) => decimal.TryParse(v?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var m) ? m : null;

    private static bool Supp(AnnualLimitDto a) => Head(a.SuppStatus, 1) is "B" or "C";

    private static bool E(string? v) => string.IsNullOrWhiteSpace(v);

    private static string? Opt(string v) => v.Length == 0 ? null : v;

    /// <summary>Trim(Mid(text, 1, n)).</summary>
    private static string Head(string? text, int n) => (text ?? "").Length > n ? text![..n].Trim() : (text ?? "").Trim();

    /// <summary>Mid(text, 1, n) (untrimmed, as SaveANN stores it).</summary>
    private static string Raw(string? text, int n) => (text ?? "").Length > n ? text![..n] : text ?? "";

    /// <summary>Trim(Mid(text, 1, IIf(InStr(text, "-") = 0, n, InStr(text, "-") - 1))).</summary>
    private static string Before(string? text, int n)
    {
        if (E(text)) return "";
        var at = text!.IndexOf('-');
        return (at < 0 ? text[..Math.Min(n, text.Length)] : text[..at]).Trim();
    }
}
