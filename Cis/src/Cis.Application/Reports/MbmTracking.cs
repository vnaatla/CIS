using System.Globalization;
using System.Reflection;
using System.Text.Json;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Reports;

/// <summary>Tracking > MBM Tracking search (FrmNewMBMTracking cmdSearch_Click / SearchRecord).</summary>
/// <param name="ReportType">CboReportType index (0 Ageband … 17 States).</param>
/// <param name="Selection">CboRptSelection text ("" for the types without choices).</param>
/// <param name="MemberType">CboMbmType: P (principal), S (supplementary), B (both).</param>
/// <param name="Filters">The criteria by name (the cis.MbmTrk_Run parameters without "@"); dates as dd/MM/yyyy, flags as true / false.</param>
public sealed record MbmTrkRequest(int ReportType, string? Selection, string? MemberType, Dictionary<string, string?> Filters);

/// <summary>Lengths.</summary>
public sealed class MbmTrkRequestValidator : Common.RequestValidator<MbmTrkRequest>
{
    /// <summary>Defines the rules.</summary>
    public MbmTrkRequestValidator()
    {
        RuleFor(r => r.ReportType).InclusiveBetween(0, 17);
        RuleFor(r => r.Selection).MaximumLength(100);
        RuleFor(r => r.Filters).Must(f => f.Count <= 200 && f.All(x => x.Key.Length <= 40 && (x.Value ?? "").Length <= 300));
    }
}

/// <summary>A report's rows (as displayed) and its totals.</summary>
public sealed record MbmTrkResult(List<string> Headers, List<List<string>> Rows, int Principals, int Supplementaries, decimal Mco, decimal Premium, decimal Access);

/// <summary>A report of the catalogue (report type, choice).</summary>
public sealed record MbmTrkReport(int Type, string Selection, bool Migrated);

/// <summary>Data access.</summary>
public interface IMbmTrkRepository
{
    /// <summary>cis.MbmTrk_Run: one case of CISMBMTracking with the typed filters.</summary>
    Task<List<Dictionary<string, object?>>> RunAsync(IDbSession session, int caseNo, IReadOnlyDictionary<string, object?> filters);

    /// <summary>cis.MbmTrk_Lookup: a combo.</summary>
    Task<List<string>> LookupAsync(IDbSession session, string kind, string? key);
}

/// <summary>Tracking > MBM Tracking (FrmNewMBMTracking).</summary>
public interface IMbmTrkService
{
    /// <summary>The report catalogue (type, choice, whether it is migrated).</summary>
    IReadOnlyList<MbmTrkReport> Reports();

    /// <summary>cmdSearch_Click.</summary>
    Task<MbmTrkResult> SearchAsync(MbmTrkRequest request);

    /// <summary>A combo: states, ins, payors (key = health type), companies (key = payor), departments (key = group company).</summary>
    Task<List<string>> LookupAsync(string kind, string? key);
}

/// <inheritdoc />
public sealed partial class MbmTrkService(IDbSessions sessions, IMbmTrkRepository repository) : IMbmTrkService
{
    /// <summary>Desktop texts.</summary>
    public const string NoReport = "Please select a report type.", NoRecord = "No record found!", BadDate = "Invalid Date Format! ";

    /// <summary>Not a desktop text: the report is in the catalogue but not ported yet.</summary>
    public const string Pending = "This report is not migrated yet.";

    private static readonly CultureInfo En = CultureInfo.InvariantCulture;

    /// <summary>A column of a block: f field, m money field, l literal, sc / sm / sp the supplementary count / MCO / premium merged in, fl field + suffix, d date,
    /// ym year / zero-padded month.</summary>
    internal sealed record Col(string K, string? N, string? V);

    internal sealed record Block(int C, bool Special, List<Col>? Cols);

    internal sealed record Helper(int C, bool Append, List<Col>? Cols);

    /// <summary>A detail listing (one row per member): the cases for P, S and Both (principals, supplementaries) and the columns of each loop.</summary>
    internal sealed record Detail(int P, int S, int Bp, int Bs, List<Col> Prin, List<Col> Supp);

    internal sealed record Entry(int Type, string Selection, string Routine, List<string> Headers, bool Special, Block? P, Block? S, Helper? Helper, Detail? Detail = null);

    internal sealed record Param(string N, string T);

    private static readonly JsonSerializerOptions Json = new() { PropertyNameCaseInsensitive = true };
    internal static readonly List<Entry> Catalogue = Load<List<Entry>>("MbmTracking.catalogue.json");
    private static readonly Dictionary<string, string> Params = Load<List<Param>>("MbmTracking.params.json").ToDictionary(p => p.N, p => p.T, StringComparer.OrdinalIgnoreCase);

    internal static T Load<T>(string name)
    {
        using var s = Assembly.GetExecutingAssembly().GetManifestResourceStream(name) ?? throw new InvalidOperationException(name + " missing");
        return JsonSerializer.Deserialize<T>(s, Json)!;
    }

    private static readonly string[] Kinds = ["states", "ins", "payors", "companies", "departments"];

    /// <inheritdoc />
    public async Task<List<string>> LookupAsync(string kind, string? key)
    {
        if (!Kinds.Contains(kind)) throw new NotFoundException("Unknown list.");
        await using var s = await sessions.OpenAsync();
        return await repository.LookupAsync(s, kind, string.IsNullOrWhiteSpace(key) ? null : key.Trim());
    }

    /// <inheritdoc />
    public IReadOnlyList<MbmTrkReport> Reports() => Catalogue.Select(e => new MbmTrkReport(e.Type, e.Selection, !e.Special || HandPorted.Contains(e.Routine))).ToList();

    /// <inheritdoc />
    public async Task<MbmTrkResult> SearchAsync(MbmTrkRequest r)
    {
        var e = Catalogue.FirstOrDefault(x => x.Type == r.ReportType && x.Selection == (r.Selection ?? "").Trim()) ?? throw new BusinessRuleException(NoReport);
        if (e.Special && !HandPorted.Contains(e.Routine)) throw new BusinessRuleException(Pending);
        var filters = Filters(r.Filters);
        // The principal type and the supplementary registration date belong to the detail listings only (their own criteria in the desktop).
        if (e.Detail is null) foreach (var k in DetailOnly) filters.Remove(k);
        var type = (r.MemberType ?? "P").Trim().ToUpperInvariant() is var t && t.Length > 0 ? t[..1] : "P";
        await using var s = await sessions.OpenAsync();
        if (e.Detail is not null) return await DetailAsync(s, e, type, filters);
        if (e.Special) return await SummaryAsync(s, e, type, filters);
        var rows = new List<List<string>>();
        int principals = 0, supps = 0;
        decimal mco = 0, premium = 0, access = 0;
        if (type is "P" or "B")
        {
            var prin = await repository.RunAsync(s, e.P!.C, filters);
            List<Dictionary<string, object?>>? supp = null;
            if (type == "B" && e.Helper is not null) supp = await repository.RunAsync(s, e.Helper.C, filters);
            var keys = e.P.Cols!.Where(c => c.K is "f" or "fe" or "d" or "fl" && !(c.N ?? "").StartsWith("Total", StringComparison.OrdinalIgnoreCase)).Select(c => c.N!)
                .Concat(e.P.Cols!.Any(c => c.K == "ym") ? ["Year", "Month"] : []).ToList();
            var merged = supp is not null && e.Helper is { Append: false }
                ? supp.GroupBy(x => Key(x, keys)).ToDictionary(g => g.Key, g => g.First()) : null;
            foreach (var x in prin)
            {
                var m = merged?.GetValueOrDefault(Key(x, keys));
                rows.Add(e.P.Cols!.Select(c => Cell(c, x, m)).ToList());
                principals += (int)N(x, "TotalP");
                mco += N(x, "MCOGross") + (m is null ? 0 : N(m, "MCOGross"));
                premium += N(x, "PremiumGross") + (m is null ? 0 : N(m, "PremiumGross"));
                access += N(x, "EAFee");
                if (m is not null) supps += (int)N(m, "TotalS");
            }
            if (supp is not null && e.Helper is { Append: true, Cols: not null })
                foreach (var x in supp)
                {
                    rows.Add(e.Helper.Cols.Select(c => Cell(c, x, null)).ToList());
                    supps += (int)N(x, "TotalS");
                    mco += N(x, "MCOGross");
                    premium += N(x, "PremiumGross");
                }
        }
        else if (e.S is { Cols: not null })
            foreach (var x in await repository.RunAsync(s, e.S.C, filters))
            {
                rows.Add(e.S.Cols.Select(c => Cell(c, x, null)).ToList());
                supps += (int)N(x, "TotalS");
                mco += N(x, "MCOGross");
                premium += N(x, "PremiumGross");
            }
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new MbmTrkResult(e.Headers, rows, principals, supps, mco, premium, access);
    }

    /// <summary>The detail listings (GeneralListing, PrintLabel, ICListing, ...): principals then supplementaries, a member (and cover id) listed once in a row.</summary>
    private async Task<MbmTrkResult> DetailAsync(IDbSession s, Entry e, string type, Dictionary<string, object?> filters)
    {
        var d = e.Detail!;
        var rows = new List<List<string>>();
        int principals = 0, supps = 0;
        decimal mco = 0, premium = 0, access = 0;
        void Add(List<Dictionary<string, object?>> source, List<Col> cols, bool supp)
        {
            var last = "";
            foreach (var x in source)
            {
                var key = T(x, "MBMNumber") + (supp ? "|" + T(x, "MBMCCoverID") : "");
                if (key == last) continue;
                last = key;
                rows.Add(cols.Select(c => DetailCell(c, x)).ToList());
                if (supp) supps++;
                else
                {
                    if (T(x, "MBMMemberType") is "C" or "S") supps++; else principals++;
                    mco += N(x, "MBMMCOFee");
                    premium += N(x, "MBMPremium");
                    access += N(x, "MBMAccessFee");
                }
            }
        }
        if (type == "S") Add(await repository.RunAsync(s, d.S, filters), d.Supp, true);
        else
        {
            var prinFilters = filters.Where(x => !SuppOnly.Contains(x.Key)).ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);
            Add(await repository.RunAsync(s, type == "B" ? d.Bp : d.P, prinFilters), d.Prin, false);
            if (type == "B") Add(await repository.RunAsync(s, d.Bs, filters), d.Supp, true);
        }
        if (rows.Count == 0) throw new NotFoundException(NoDetailRecord);
        return new MbmTrkResult(e.Headers, rows, principals, supps, mco, premium, access);
    }

    private static readonly string[] DetailOnly = ["detailPrinType", "suppRegFrom", "suppRegTo", "suppRegBlank"];
    private static readonly string[] SuppOnly = ["suppRegFrom", "suppRegTo", "suppRegBlank"];

    /// <summary>Desktop text of the detail listings.</summary>
    public const string NoDetailRecord = "Sorry, No Record Found";

    private static string DetailCell(Col c, IReadOnlyDictionary<string, object?> x) => c.K switch
    {
        "cv" => T(x, c.N!) is { Length: > 0 } v ? v : c.V ?? "",
        "nm" => T(x, "MBMName") + (T(x, "MBMStatus") switch { "S" => " (Suspended)", "C" => " (Cancelled)", "D" => " (Deleted)", _ => "" }),
        "d" => x.GetValueOrDefault(c.N!) is DateTime dt ? dt.ToString("dd/MM/yyyy", En) : T(x, c.N!),
        _ => Cell(c, x, null),
    };

    /// <summary>The typed filters: only known names, each converted to its parameter type.</summary>
    internal static Dictionary<string, object?> Filters(Dictionary<string, string?> input) => Filters(input, Params);

    /// <summary>The typed filters for a given parameter list (name -> SQL type).</summary>
    internal static Dictionary<string, object?> Filters(Dictionary<string, string?> input, IReadOnlyDictionary<string, string> types)
    {
        var p = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, raw) in input)
        {
            if (!types.TryGetValue(k, out var type) || string.IsNullOrWhiteSpace(raw)) continue;
            var v = raw.Trim();
            p[k] = type switch
            {
                "bit" => bool.TryParse(v, out var b) ? b : v == "1",
                "datetime" => ClinicInvoiceRules.Date(v) ?? throw new BusinessRuleException(BadDate),
                "money" => decimal.TryParse(v, NumberStyles.Number, En, out var d) ? d : throw new BusinessRuleException("Invalid amount."),
                _ => v,
            };
        }
        return p;
    }

    private static string Key(IReadOnlyDictionary<string, object?> r, List<string> keys) => string.Join("\u0001", keys.Select(k => T(r, k)));

    internal static string Cell(Col c, IReadOnlyDictionary<string, object?> x, IReadOnlyDictionary<string, object?>? supp) => c.K switch
    {
        "fe" => T(x, c.N!) is { Length: > 0 } v ? v : "Empty",
        "mp" => x.GetValueOrDefault(c.N!) is null or DBNull ? "" : N(x, c.N!).ToString("#,##0.00;(#,##0.00)", En),
        "avg" => N(x, c.V!) == 0 ? "" : Math.Round(N(x, c.N!) / N(x, c.V!), 2, MidpointRounding.ToEven).ToString("#,##0.00", En),
        "pct" => N(x, c.V!) == 0 ? "" : Math.Round(N(x, c.N!) / N(x, c.V!) * 100, 0, MidpointRounding.ToEven).ToString(En),
        "cat" => T(x, c.N!.Split('|')[0]) + c.V + T(x, c.N!.Split('|')[1]),
        "f" => T(x, c.N!),
        "m" => x.GetValueOrDefault(c.N!) is null or DBNull ? "0.00" : N(x, c.N!).ToString("#,##0.00", En),
        "l" => c.V ?? "",
        "fl" => T(x, c.N!) + c.V,
        "d" => x.GetValueOrDefault(c.N!) is DateTime d ? d.ToString("dd/MM/yyyy", En) : T(x, c.N!),
        "ym" => T(x, "Year") + "/" + (T(x, "Month").Length == 1 ? "0" + T(x, "Month") : T(x, "Month")),
        "sc" => supp is null ? "0" : T(supp, "TotalS"),
        "sm" => supp is null ? "0.00" : N(supp, "MCOGross").ToString("#,##0.00", En),
        "sp" => supp is null ? "0.00" : N(supp, "PremiumGross").ToString("#,##0.00", En),
        _ => "",
    };

    internal static string T(IReadOnlyDictionary<string, object?> r, string k) => r.GetValueOrDefault(k) switch
    {
        null or DBNull => "",
        DateTime d => d.ToString("dd/MM/yyyy", En),
        var o => Convert.ToString(o, En)!.Trim(),
    };

    internal static decimal N(IReadOnlyDictionary<string, object?> r, string k) => decimal.TryParse(T(r, k), NumberStyles.Any, En, out var d) ? d : 0;
}
