using Cis.Application.Abstractions;
using Cis.Domain.Errors;
using FluentValidation;
using static Cis.Application.Reports.MbmTrkService;

namespace Cis.Application.Reports;

/// <summary>Claim / Bordx > Claim Tracking search (frmClaimRpt CmdSearch_Click / SearchRecords).</summary>
/// <param name="ReportType">CboReportType index (0 Admission … 23 Verification of Claims).</param>
/// <param name="Selection">CboSelection text ("" for the types without choices).</param>
/// <param name="Filters">The criteria by name (the cis.ClmTrk_Run parameters without "@"); dates as dd/MM/yyyy, flags as true / false.</param>
public sealed record ClmTrkRequest(int ReportType, string? Selection, Dictionary<string, string?> Filters);

/// <summary>Lengths.</summary>
public sealed class ClmTrkRequestValidator : Common.RequestValidator<ClmTrkRequest>
{
    /// <summary>Defines the rules.</summary>
    public ClmTrkRequestValidator()
    {
        RuleFor(r => r.ReportType).InclusiveBetween(0, 23);
        RuleFor(r => r.Selection).MaximumLength(100);
        RuleFor(r => r.Filters).Must(f => f.Count <= 200 && f.All(x => x.Key.Length <= 40 && (x.Value ?? "").Length <= 300));
    }
}

/// <summary>A claim tracking report's rows (as displayed) and the record count the desktop shows.</summary>
public sealed record ClmTrkResult(List<string> Headers, List<List<string>> Rows, int Records);

/// <summary>Data access.</summary>
public interface IClmTrkRepository
{
    /// <summary>cis.ClmTrk_Run: one case of CISClaimTracking with the typed filters.</summary>
    /// <param name="session">The session.</param>
    /// <param name="caseNo">The case.</param>
    /// <param name="filters">The typed filters.</param>
    /// <param name="proc">cis.ClmTrk_Run (CISClaimTracking) or cis.ClmTrkSimple_Run (CISClaimTracking_Simple).</param>
    Task<List<Dictionary<string, object?>>> RunAsync(IDbSession session, int caseNo, IReadOnlyDictionary<string, object?> filters, string proc = "cis.ClmTrk_Run");

    /// <summary>cis.ClmTrk_Lookup: a combo.</summary>
    Task<List<string>> LookupAsync(IDbSession session, string kind, string? key);
}

/// <summary>Claim / Bordx > Claim Tracking (frmClaimRpt).</summary>
public interface IClmTrkService
{
    /// <summary>The report catalogue (type, choice, whether it is migrated).</summary>
    /// <param name="variant">main, simple, jupiter, jupitersimple or sun.</param>
    IReadOnlyList<MbmTrkReport> Reports(string variant = "main");

    /// <summary>CmdSearch_Click.</summary>
    /// <param name="request">The search.</param>
    /// <param name="variant">main, simple, jupiter, jupitersimple or sun.</param>
    Task<ClmTrkResult> SearchAsync(ClmTrkRequest request, string variant = "main");

    /// <summary>A combo: polstatus, states, clinicgroups, clinics, payors (key = health type), companies (key = payor), departments (key = group company).</summary>
    Task<List<string>> LookupAsync(string kind, string? key);
}

/// <inheritdoc />
public sealed partial class ClmTrkService(IDbSessions sessions, IClmTrkRepository repository) : IClmTrkService
{
    internal static readonly List<Entry> Catalogue = Load<List<Entry>>("ClmTracking.catalogue.json");

    /// <summary>The catalogue of each variant form (generated from that form by sql/tools/analyse-clmtrk.js).</summary>
    internal static readonly Dictionary<string, List<Entry>> Catalogues = new()
    {
        ["main"] = Catalogue,
        ["simple"] = Load<List<Entry>>("ClmTrackingSimple.catalogue.json"),
        ["jupiter"] = Load<List<Entry>>("ClmTrackingJupiter.catalogue.json"),
        ["jupitersimple"] = Load<List<Entry>>("ClmTrackingJupiterSimple.catalogue.json"),
        ["sun"] = Load<List<Entry>>("ClmTrackingSun.catalogue.json"),
    };

    /// <summary>Per variant: its procedure, the last report type its combo offers, and the hand-ported listings whose code differs in that form and are not ported yet.</summary>
    private static readonly Dictionary<string, (string Proc, int MaxType, string[] Pending)> Variants = new()
    {
        ["main"] = ("cis.ClmTrk_Run", 23, []),
        ["simple"] = ("cis.ClmTrkSimple_Run", 19, []),
        ["jupiter"] = ("cis.ClmTrk_Run", 20, []),
        ["jupitersimple"] = ("cis.ClmTrkSimple_Run", 18, []),
        ["sun"] = ("cis.ClmTrk_Run", 22, []),
    };

    /// <summary>The procedure of the variant being searched (the service is scoped to one request).</summary>
    private string proc = "cis.ClmTrk_Run";

    /// <summary>The variant being searched (for the listings whose columns differ by form).</summary>
    private string variant = "main";

    private Task<List<Dictionary<string, object?>>> RunAsync(IDbSession s, int caseNo, IReadOnlyDictionary<string, object?> filters) => repository.RunAsync(s, caseNo, filters, proc);

    private static (string Proc, int MaxType, string[] Pending) VariantOf(string variant) =>
        Variants.TryGetValue(variant, out var v) ? v : throw new NotFoundException("Unknown tracking screen.");

    private static bool Runs(Entry e, string variant) => (!e.Special || IsHandPorted(e.Routine)) && !VariantOf(variant).Pending.Contains(e.Routine);
    private static readonly Dictionary<string, string> Params = Load<List<Param>>("ClmTracking.params.json").ToDictionary(p => p.N, p => p.T, StringComparer.OrdinalIgnoreCase);

    private static readonly string[] Kinds = ["polstatus", "states", "clinicgroups", "clinics", "payors", "companies", "departments"];

    /// <inheritdoc />
    public async Task<List<string>> LookupAsync(string kind, string? key)
    {
        if (!Kinds.Contains(kind)) throw new NotFoundException("Unknown list.");
        await using var s = await sessions.OpenAsync();
        return await repository.LookupAsync(s, kind, string.IsNullOrWhiteSpace(key) ? null : key.Trim());
    }

    /// <inheritdoc />
    public IReadOnlyList<MbmTrkReport> Reports(string variant = "main") =>
        Catalogues[variant is var v && Variants.ContainsKey(v) ? v : throw new NotFoundException("Unknown tracking screen.")]
            .Where(e => e.Type <= Variants[variant].MaxType).Select(e => new MbmTrkReport(e.Type, e.Selection, Runs(e, variant))).ToList();

    /// <summary>The typed filters (only the known names).</summary>
    internal static Dictionary<string, object?> Filters(Dictionary<string, string?> input) => MbmTrkService.Filters(input, Params);

    /// <inheritdoc />
    public async Task<ClmTrkResult> SearchAsync(ClmTrkRequest r, string variant = "main")
    {
        var vv = VariantOf(variant);
        proc = vv.Proc;
        this.variant = variant;
        if (r.ReportType > vv.MaxType) throw new BusinessRuleException(NoReport);
        var e = Catalogues[variant].FirstOrDefault(x => x.Type == r.ReportType && x.Selection == (r.Selection ?? "").Trim()) ?? throw new BusinessRuleException(NoReport);
        if (!Runs(e, variant)) throw new BusinessRuleException(Pending);
        var filters = Filters(r.Filters);
        await using var s = await sessions.OpenAsync();
        if (e.Special)
        {
            var special = await SpecialRowsAsync(s, e, filters);
            if (special.Count == 0) throw new NotFoundException(NoRecord);
            return new ClmTrkResult(e.Headers, special, special.Count);
        }
        // The treatment reason descriptions (the desktop looked each row's code up with CISClaimTracking case 10).
        var reasons = e.P!.Cols!.Any(c => c.K == "tr")
            ? (await repository.LookupAsync(s, "reasons", null)).Select(x => x.Split('|', 2)).GroupBy(x => x[0]).ToDictionary(g => g.Key, g => g.First()[1])
            : null;
        string Value(Col c, IReadOnlyDictionary<string, object?> x) => c.K == "tr" ? reasons!.GetValueOrDefault(T(x, c.N!)) ?? "" : Cell(c, x, null);
        var rows = (await RunAsync(s, e.P.C, filters)).Select(x => e.P.Cols!.Select(c => Value(c, x)).ToList()).ToList();
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new ClmTrkResult(e.Headers, rows, rows.Count);
    }
}
