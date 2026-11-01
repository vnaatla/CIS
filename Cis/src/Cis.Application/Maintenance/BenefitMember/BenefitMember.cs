using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Documents;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.BenefitMember;

/// <summary>A row of the Benefit Entry grid (FGrid: Cat Code, Calculation, No of Visit Per day / Per Annum, RM Per Visit, Percentage).</summary>
/// <param name="BnfCode">Category code (BNFCode).</param>
/// <param name="Calculation">1 - Lower of Limit or % · 2 - Higher of Limit or % · 3 - Ascharge · 4 - BNF Limit.</param>
/// <param name="NoOfDays">No of Visit Per day.</param>
/// <param name="NoOfAnn">No of Visit Per Annum.</param>
/// <param name="RmPerVisit">RM Per Visit.</param>
/// <param name="Percentage">Percentage.</param>
public record BenefitEntryRow(string BnfCode, string? Calculation, string? NoOfDays, string? NoOfAnn, string? RmPerVisit, string? Percentage);

/// <summary>Combos of the Benefit Entry tab.</summary>
/// <param name="HealthTypes">Fixed: S - Sihat, N - Non-Sihat.</param>
/// <param name="InsuredTypes">INS.</param>
/// <param name="Plans">PLN codes.</param>
/// <param name="Categories">BNFCategoryNew (the grid rows).</param>
public record BenefitEntrySetup(List<LookupOption> HealthTypes, List<LookupOption> InsuredTypes, List<LookupOption> Plans, List<LookupOption> Categories);

/// <summary>Combos of the Benefit Listing tab.</summary>
/// <param name="HealthTypes">HLT.</param>
/// <param name="InsuredTypes">INS.</param>
/// <param name="Plans">PLN codes.</param>
/// <param name="Payors">HISDB PAY.</param>
/// <param name="GroupCompanies">Distinct HISDB PLN.GRPCompany.</param>
/// <param name="Categories">BNFCategoryNew.</param>
public record BenefitListingSetup(List<LookupOption> HealthTypes, List<LookupOption> InsuredTypes, List<LookupOption> Plans, List<LookupOption> Payors,
    List<LookupOption> GroupCompanies, List<LookupOption> Categories);

/// <summary>Benefit Entry key (Health Type, Insured Type, Plan Code, Eff Date as typed dd/mm/yyyy).</summary>
/// <param name="HealthType">Health type (first character is stored).</param>
/// <param name="InsuredType">Insured type (first character).</param>
/// <param name="PlanCode">Plan code (first 4 characters).</param>
/// <param name="EffDate">Effective date dd/mm/yyyy.</param>
public record BenefitEntryKey(string? HealthType, string? InsuredType, string? PlanCode, string? EffDate);

/// <summary>Save of the Benefit Entry tab.</summary>
/// <param name="HealthType">Health type.</param>
/// <param name="InsuredType">Insured type.</param>
/// <param name="PlanCode">Plan code.</param>
/// <param name="EffDate">Effective date dd/mm/yyyy.</param>
/// <param name="Rows">Grid rows.</param>
public record BenefitEntrySaveRequest(string? HealthType, string? InsuredType, string? PlanCode, string? EffDate, List<BenefitEntryRow>? Rows);

/// <summary>Filters of the Benefit Listing tab (all optional; codes as chosen in the combos).</summary>
/// <param name="BnfCode">Category.</param>
/// <param name="HealthType">Health type.</param>
/// <param name="InsuredType">Insured type.</param>
/// <param name="PlanCode">Plan code.</param>
/// <param name="EffDate">Eff Date dd/mm/yyyy (ignored when not a valid date).</param>
/// <param name="PayorCode">Payor code.</param>
/// <param name="GroupCompany">Group company (prefix search).</param>
/// <param name="Calculation">BNF Calculation.</param>
/// <param name="NoOfDays">No of Days.</param>
/// <param name="RmPerVisit">RM Per Visit.</param>
/// <param name="NoOfAnn">No of Ann.</param>
/// <param name="Percentage">BNF Percentage.</param>
public record BenefitListingRequest(string? BnfCode, string? HealthType, string? InsuredType, string? PlanCode, string? EffDate, string? PayorCode,
    string? GroupCompany, string? Calculation, string? NoOfDays, string? RmPerVisit, string? NoOfAnn, string? Percentage);

/// <summary>Column lengths of BNFNew; the desktop's own checks (and their order) are in <see cref="BenefitMemberService"/>.</summary>
public sealed class BenefitEntrySaveRequestValidator : RequestValidator<BenefitEntrySaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public BenefitEntrySaveRequestValidator()
    {
        RuleForEach(r => r.Rows).ChildRules(row =>
        {
            row.RuleFor(x => x.BnfCode).NotEmpty().MaximumLength(10);
            row.RuleFor(x => x.RmPerVisit).MaximumLength(15).WithMessage("RM Per Visit must be at most 15 characters.");
            row.RuleFor(x => x.Percentage).MaximumLength(15).WithMessage("Percentage must be at most 15 characters.");
            row.RuleFor(x => x.NoOfDays).Must(BenefitMemberService.IsSmallIntOrEmpty).WithMessage("No of Visit Per day must be a whole number.");
            row.RuleFor(x => x.NoOfAnn).Must(BenefitMemberService.IsSmallIntOrEmpty).WithMessage("No of Visit Per Annum must be a whole number.");
        });
    }
}

/// <summary>Benefit Listing filters: lengths only.</summary>
public sealed class BenefitListingRequestValidator : RequestValidator<BenefitListingRequest>
{
    /// <summary>Defines the rules.</summary>
    public BenefitListingRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(50);
        RuleFor(r => r.RmPerVisit).MaximumLength(15);
        RuleFor(r => r.Percentage).MaximumLength(15);
    }
}

/// <summary>Data access of the Benefit Entry and Benefit Listing tabs.</summary>
public interface IBenefitMemberRepository
{
    /// <summary>INS, PLN, categories.</summary>
    Task<(List<LookupOption> Insured, List<LookupOption> Plans, List<LookupOption> Categories)> EntrySetupAsync(IDbSession session);

    /// <summary>Listing combos.</summary>
    Task<BenefitListingSetup> ListingSetupAsync(IDbSession session);

    /// <summary>BNFNew rows of the key.</summary>
    Task<List<BenefitEntryRow>> EntrySearchAsync(IDbSession session, string plnCode, string insCode, string hltCode, DateTime effDate);

    /// <summary>True when plan + insured type + date already has rows.</summary>
    Task<bool> EntryExistsAsync(IDbSession session, string plnCode, string insCode, DateTime effDate);

    /// <summary>Inserts one BNFNew row.</summary>
    Task EntryInsertAsync(IDbSession session, string plnCode, string insCode, string hltCode, DateTime effDate, BenefitEntryRow row, string userCode, DateTime today);

    /// <summary>VIEW_CISBNF_PLN_SubCat rows of the filters (already reduced to the desktop's SQL values).</summary>
    Task<List<Dictionary<string, object?>>> ListingAsync(IDbSession session, BenefitListingRequest filters, DateTime? effDate);
}

/// <summary>Maintenance > Benefit > Member, Benefit Entry and Benefit Listing tabs.</summary>
public interface IBenefitMemberService
{
    /// <summary>BNFcombolisting.</summary>
    Task<BenefitEntrySetup> EntrySetupAsync();

    /// <summary>searchtBNF: all four keys needed; returns the saved rows (none = empty grid, no message, as in the desktop).</summary>
    Task<List<BenefitEntryRow>> EntrySearchAsync(BenefitEntryKey key);

    /// <summary>SaveBNF → "Record Saved...".</summary>
    Task<string> EntrySaveAsync(BenefitEntrySaveRequest request, string userCode);

    /// <summary>BNFCombolisting2.</summary>
    Task<BenefitListingSetup> ListingSetupAsync();

    /// <summary>searchBNF → BNFListing; "Record not found" when empty.</summary>
    Task<List<Dictionary<string, object?>>> ListingAsync(BenefitListingRequest filters);

    /// <summary>cmdprint_Click on Benefit Listing: the list as Excel ("Report cannot be printed without any record.").</summary>
    Task<FileOutput> ListingExportAsync(BenefitListingRequest filters);
}

/// <inheritdoc />
public sealed class BenefitMemberService(IDbSessions sessions, IBenefitMemberRepository repository, IReportExporter exporter, TimeProvider clock)
    : IBenefitMemberService
{
    /// <summary>Desktop texts (spelling and spacing kept).</summary>
    public const string SearchHealth = "Please key in Health Type !", SearchInsured = "Please key in Insured Type !", SearchPlan = "Please key in Plan Code!",
        SearchDate = "Please key in Effective Date!", SearchValidDate = "Please key in VALID Effective Date!", SaveHealth = "Please Enter Health Type !",
        SaveInsured = "Please Enter Insured Type!", SavePlan = "Please Enter Plan Code!", SaveDate = "Please Enter (Valid )effectice Date!",
        Duplicated = "Record Duplicated! Please Change The Corresponding Field", Saved = "Record Saved...", NotFound = "Record not found",
        NothingToPrint = "Report cannot be printed without any record.";

    /// <summary>lstBNFList columns (the Excel export header).</summary>
    public static readonly IReadOnlyList<ReportColumn> ListingColumns =
    [
        new("BNFCode", "BNF Code"), new("BNFCATDescription", "Description"), new("HLTCode", "HLT Code"), new("INSCode", "INS Code"),
        new("PLNCode", "PLN Code"), new("BNFCalculation", "BNF Calculation"), new("BNFNoOfDays", "No of Days", ReportFormat.Number),
        new("BNFNoOfAnn", "No of Annum", ReportFormat.Number), new("BNFRMPerVisit", "RM Per Visit"), new("BNFPercentage", "BNF Percentage"),
        new("BNFEffDate", "Eff Date", ReportFormat.Date), new("PAYCode", "Payor Code"), new("GRPCompany", "Group Company"),
    ];

    /// <inheritdoc />
    public async Task<BenefitEntrySetup> EntrySetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        var (insured, plans, categories) = await repository.EntrySetupAsync(s);
        return new BenefitEntrySetup([new("S", "Sihat"), new("N", "Non-Sihat")], insured, plans, categories);
    }

    /// <inheritdoc />
    public async Task<List<BenefitEntryRow>> EntrySearchAsync(BenefitEntryKey key)
    {
        if (Blank(key.HealthType)) throw new BusinessRuleException(SearchHealth);
        if (Blank(key.InsuredType)) throw new BusinessRuleException(SearchInsured);
        if (Blank(key.PlanCode)) throw new BusinessRuleException(SearchPlan);
        if (Blank(key.EffDate)) throw new BusinessRuleException(SearchDate);
        var date = ParseDate(key.EffDate) ?? throw new BusinessRuleException(SearchValidDate);
        await using var s = await sessions.OpenAsync();
        return await repository.EntrySearchAsync(s, Plan(key.PlanCode), First(key.InsuredType), First(key.HealthType), date);
    }

    /// <inheritdoc />
    public async Task<string> EntrySaveAsync(BenefitEntrySaveRequest request, string userCode)
    {
        if (Blank(request.HealthType)) throw new BusinessRuleException(SaveHealth);
        if (Blank(request.InsuredType)) throw new BusinessRuleException(SaveInsured);
        if (Blank(request.PlanCode)) throw new BusinessRuleException(SavePlan);
        var date = ParseDate(request.EffDate) ?? throw new BusinessRuleException(SaveDate);
        var (pln, ins, hlt) = (Plan(request.PlanCode), First(request.InsuredType), First(request.HealthType));
        await using var s = await sessions.BeginAsync();
        if (await repository.EntryExistsAsync(s, pln, ins, date)) throw new ConflictException(Duplicated);
        var today = clock.GetLocalNow().Date;
        foreach (var row in (request.Rows ?? []).Select(Clean).Where(Qualifies))
            await repository.EntryInsertAsync(s, pln, ins, hlt, date, row, userCode, today);
        await s.CommitAsync();
        return Saved;
    }

    /// <inheritdoc />
    public async Task<BenefitListingSetup> ListingSetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.ListingSetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> ListingAsync(BenefitListingRequest filters)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.ListingAsync(s, Reduce(filters), ParseDate(filters.EffDate));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<FileOutput> ListingExportAsync(BenefitListingRequest filters)
    {
        List<Dictionary<string, object?>> rows;
        await using (var s = await sessions.OpenAsync()) rows = await repository.ListingAsync(s, Reduce(filters), ParseDate(filters.EffDate));
        if (rows.Count == 0) throw new BusinessRuleException(NothingToPrint);
        return exporter.Export(new ReportDocument("Benefit Listing", ListingColumns, rows) { FileName = "BenefitListing" }, ReportFileType.Excel);
    }

    /// <summary>
    /// Which grid rows SaveBNF writes: Calculation 1 or 2 with RM Per Visit or Percentage, 3 always, 4 with RM Per Visit; other rows
    /// (no calculation) are skipped.
    /// </summary>
    public static bool Qualifies(BenefitEntryRow r) => r.Calculation switch
    {
        "1" or "2" => !string.IsNullOrEmpty(r.RmPerVisit) || !string.IsNullOrEmpty(r.Percentage),
        "3" => true,
        "4" => !string.IsNullOrEmpty(r.RmPerVisit),
        _ => false,
    };

    /// <summary>FGrid_Click rules (Calculation 1–4, numeric day / annum / percentage, else blank) and ChkStr (UCase(Trim)); RM Per Visit is ""
    /// rather than NULL when empty, as the desktop wrote it.</summary>
    public static BenefitEntryRow Clean(BenefitEntryRow r)
    {
        static string? Num(string? v) => decimal.TryParse(v?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out _) ? v!.Trim() : null;
        var calc = Num(r.Calculation);
        if (calc is not null && (!decimal.TryParse(calc, NumberStyles.Number, CultureInfo.InvariantCulture, out var c) || c < 1 || c > 4)) calc = null;
        return new(r.BnfCode.Trim(), calc, Num(r.NoOfDays), Num(r.NoOfAnn), (r.RmPerVisit ?? "").Trim().ToUpperInvariant(),
            Num(r.Percentage)?.ToUpperInvariant());
    }

    /// <summary>True for empty or a whole number that fits smallint (BNFNoOfDays / BNFNoOfAnn).</summary>
    public static bool IsSmallIntOrEmpty(string? v) => string.IsNullOrWhiteSpace(v) ||
        !decimal.TryParse(v.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out _) ||
        short.TryParse(v.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out _);

    /// <summary>searchBNF's SQL values: category before "--", first character of health / insured, first 4 of the plan, payor and group
    /// before "-" (else first 4 characters), the rest trimmed.</summary>
    public static BenefitListingRequest Reduce(BenefitListingRequest f) => new(
        Before(f.BnfCode, "--", 10), Opt(f.HealthType) is { } h ? h[..1] : null, Opt(f.InsuredType) is { } i ? i[..1] : null,
        Opt(f.PlanCode) is { } p ? Plan(p) : null, f.EffDate, Before(f.PayorCode, "-", 4), Before(f.GroupCompany, "-", 4), Opt(f.Calculation),
        Opt(f.NoOfDays), Opt(f.RmPerVisit), Opt(f.NoOfAnn), Opt(f.Percentage));

    /// <summary>The masked dd/mm/yyyy box (IsDate); ISO accepted too.</summary>
    public static DateTime? ParseDate(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static string? Before(string? v, string mark, int fallback)
    {
        if (Opt(v) is not { } t) return null;
        var at = t.IndexOf(mark, StringComparison.Ordinal);
        return (at < 0 ? t[..Math.Min(fallback, t.Length)] : t[..at]).Trim();
    }

    private static string Plan(string? v) => v!.Trim()[..Math.Min(4, v.Trim().Length)].Trim();

    private static string First(string? v) => v!.Trim()[..1].ToUpperInvariant();

    private static bool Blank(string? v) => string.IsNullOrWhiteSpace(v);

    private static string? Opt(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
}
