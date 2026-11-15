using Cis.Application.Abstractions;
using Cis.Domain.Errors;

namespace Cis.Application.Reports;

/// <summary>The member tracking summaries that combine several CISMBMTracking cases (ported by hand from FrmNewMBMTracking).</summary>
public sealed partial class MbmTrkService
{
    /// <summary>A category column pair (amount, count) of the INS Type &amp; MBM Type summaries: its case, extra criteria and whether it counts supplementaries.</summary>
    private sealed record Cat(int Case, string? Ins, string? AdultChild, string Count, string Amount, bool Supp);

    private sealed record Pivot(int BaseP, int BaseS, string[] Keys, Cat[] Cats, bool Premium);

    private static Cat C(int c, string? ins, string? ac, string count, string amount, bool supp = false) => new(c, ins, ac, count, amount, supp);

    /// <summary>BordxMCOSum, BordxPremiumSum, InvoiceMCOSum, InvoicePremiumSum: a row per base group, then per category its amount and count, totals.</summary>
    private static readonly Dictionary<string, Pivot> Pivots = new()
    {
        ["BordxMCOSum"] = new(37, 38, ["PAYCode", "MBMBordxDate"], [C(39, null, null, "TotalIC", "MCOFEE"), C(40, null, null, "TotalIA", "MCOFEE"),
            C(41, null, null, "TotalF", "MCOFEE"), C(42, null, null, "TotalGC", "MCOFEE"), C(43, null, null, "TotalGA", "MCOFEE"), C(44, null, null, "TotalH", "MCOFEE"),
            C(45, null, null, "TotalFS", "MCOFEE", true), C(46, null, null, "TotalHS", "MCOFEE", true)], false),
        ["BordxPremiumSum"] = new(47, 56, ["PAYCode", "MBMBordxDate"], [C(48, null, null, "TotalIC", "Premium"), C(49, null, null, "TotalIA", "Premium"),
            C(50, null, null, "TotalF", "Premium"), C(51, null, null, "TotalFS", "Premium", true), C(52, null, null, "TotalGC", "Premium"),
            C(53, null, null, "TotalGA", "Premium"), C(54, null, null, "TotalH", "Premium"), C(55, null, null, "TotalHS", "Premium", true)], true),
        ["InvoiceMCOSum"] = new(130, 131, ["PAYCode", "MBMInvoiceNo", "MBMInvoiceDate"], [C(132, "I", "C", "TotalMBM", "MCOFEE"), C(132, "I", "A", "TotalMBM", "MCOFEE"),
            C(132, "F", null, "TotalMBM", "MCOFEE"), C(132, "G", "C", "TotalMBM", "MCOFEE"), C(132, "G", "A", "TotalMBM", "MCOFEE"), C(132, "H", null, "TotalMBM", "MCOFEE"),
            C(133, "F", null, "TotalMBMS", "MCOFEE", true), C(133, "H", null, "TotalMBMS", "MCOFEE", true)], false),
        ["InvoicePremiumSum"] = new(130, 131, ["PAYCode", "MBMInvoiceNo", "MBMInvoiceDate"], [C(134, "I", "C", "TotalMBM", "Premium"),
            C(134, "I", "A", "TotalMBM", "Premium"), C(134, "F", null, "TotalMBM", "Premium"), C(135, "F", null, "TotalMBMS", "Premium", true),
            C(134, "G", "C", "TotalMBM", "Premium"), C(134, "G", "A", "TotalMBM", "Premium"), C(134, "H", null, "TotalMBM", "Premium"),
            C(135, "H", null, "TotalMBMS", "Premium", true)], true),
    };

    /// <summary>The summaries ported by hand.</summary>
    internal static readonly string[] HandPorted = [.. Pivots.Keys, "MCOInvoiceSum", "EffPeriodSum"];

    private Task<MbmTrkResult> SummaryAsync(IDbSession s, Entry e, string type, Dictionary<string, object?> filters) => e.Routine switch
    {
        "MCOInvoiceSum" => McoInvoiceSumAsync(s, e, type, filters),
        "EffPeriodSum" => EffPeriodSumAsync(s, e, type, filters),
        _ => PivotAsync(s, e, Pivots[e.Routine], type, filters),
    };

    private static string Key(IReadOnlyDictionary<string, object?> r, string[] keys) => string.Join("\u0001", keys.Select(k => T(r, k)));

    private static string Empty(string v) => v.Length == 0 ? "Empty" : v;

    private async Task<MbmTrkResult> PivotAsync(IDbSession s, Entry e, Pivot p, string type, Dictionary<string, object?> filters)
    {
        var groups = await repository.RunAsync(s, type == "S" ? p.BaseS : p.BaseP, filters);
        var sums = new List<Dictionary<string, (decimal Count, decimal Amount)>>();
        foreach (var c in p.Cats)
        {
            var run = c.Supp ? type is "B" or "S" : type is "P" or "B";
            var by = new Dictionary<string, (decimal, decimal)>();
            if (run)
            {
                var f = new Dictionary<string, object?>(filters, StringComparer.OrdinalIgnoreCase);
                if (c.Ins is not null) f["ins"] = c.Ins;
                if (c.AdultChild is not null) f["adultChild"] = c.AdultChild;
                // A category's own criteria replace the user's insured type / adult-child ones, as the desktop appended both (ANDed: the same rows unless they differ).
                if (c.Ins is not null && filters.TryGetValue("ins", out var ui) && !Equals(ui, c.Ins)) f["ins"] = "\u0000";
                if (c.AdultChild is not null && filters.TryGetValue("adultChild", out var ua) && !Equals(ua, c.AdultChild)) f["adultChild"] = "\u0000";
                foreach (var x in await repository.RunAsync(s, c.Case, f))
                {
                    var k = Key(x, p.Keys);
                    var (n, a) = by.GetValueOrDefault(k);
                    by[k] = (n + N(x, c.Count), a + N(x, c.Amount));
                }
            }
            sums.Add(by);
        }
        var rows = new List<List<string>>();
        int principals = 0, supps = 0;
        decimal amount = 0;
        foreach (var g in groups)
        {
            var k = Key(g, p.Keys);
            var vals = sums.Select(d => d.GetValueOrDefault(k)).ToList();
            var row = p.Keys.Select(x => Empty(T(g, x))).ToList();
            row.AddRange(vals.Select(v => v.Amount.ToString("#,##0.00", En)));
            row.Add(vals.Sum(v => v.Amount).ToString("#,##0.00", En));
            row.AddRange(vals.Select(v => ((int)v.Count).ToString(En)));
            row.Add(((int)vals.Sum(v => v.Count)).ToString(En));
            rows.Add(row);
            for (var i = 0; i < p.Cats.Length; i++)
                if (p.Cats[i].Supp) supps += (int)vals[i].Count; else principals += (int)vals[i].Count;
            amount += vals.Sum(v => v.Amount);
        }
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new MbmTrkResult(e.Headers, rows, principals, supps, p.Premium ? 0 : amount, p.Premium ? amount : 0, 0);
    }

    /// <summary>MCOInvoiceSum: Both from case 178 (principal / dependent pivot), P from case 132, S from case 133.</summary>
    private async Task<MbmTrkResult> McoInvoiceSumAsync(IDbSession s, Entry e, string type, Dictionary<string, object?> filters)
    {
        var rows = new List<List<string>>();
        int principals = 0, supps = 0;
        decimal mco = 0;
        string Inv(IReadOnlyDictionary<string, object?> x) => Empty(T(x, "MBMInvoiceNo"));
        string Date(IReadOnlyDictionary<string, object?> x) => Empty(T(x, "MBMInvoiceDate"));
        string M(IReadOnlyDictionary<string, object?> x, string k) => N(x, k).ToString("#,##0.00", En);
        foreach (var x in await repository.RunAsync(s, type switch { "B" => 178, "S" => 133, _ => 132 }, filters))
        {
            if (type == "B")
            {
                rows.Add([T(x, "PAYCode"), Inv(x), Date(x), T(x, "Principal_mbmcnt"), T(x, "Dependent_mbmcnt"), M(x, "Principal_amt"), M(x, "Dependent_amt"),
                    T(x, "MBMGroupCompany"), T(x, "Subsidiary"), T(x, "OPAccCode")]);
                principals += (int)N(x, "Principal_mbmcnt");
                supps += (int)N(x, "Dependent_mbmcnt");
                mco += N(x, "Principal_amt") + N(x, "Dependent_amt");
            }
            else if (type == "S")
            {
                rows.Add([T(x, "PAYCode"), Inv(x), Date(x), "0", T(x, "TotalMBMS"), "0.00", M(x, "MCOFEE"), "", "", ""]);
                supps += (int)N(x, "TotalMBMS");
                mco += N(x, "MCOFEE");
            }
            else
            {
                rows.Add([T(x, "PAYCode"), Inv(x), Date(x), T(x, "TotalMBM"), "", M(x, "MCOFEE"), "", T(x, "MBMGroupCompany"), T(x, "Subsidiary"), ""]);
                principals += (int)N(x, "TotalMBM");
                mco += N(x, "MCOFEE");
            }
        }
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new MbmTrkResult(e.Headers, rows, principals, supps, mco, 0, 0);
    }

    /// <summary>EffPeriodSum: effective days (days insured + 1), principals (case 80) and supplementaries (case 81) per day.</summary>
    private async Task<MbmTrkResult> EffPeriodSumAsync(IDbSession s, Entry e, string type, Dictionary<string, object?> filters)
    {
        static string Day(IReadOnlyDictionary<string, object?> x) => int.TryParse(T(x, "DayInsured"), out var d) ? (d + 1).ToString(En) : T(x, "DayInsured");
        var rows = new List<List<string>>();
        int principals = 0, supps = 0;
        if (type == "S")
            foreach (var x in await repository.RunAsync(s, 81, filters))
            {
                rows.Add([Day(x), "0", T(x, "TotalS")]);
                supps += (int)N(x, "TotalS");
            }
        else
        {
            var supp = type == "B" ? (await repository.RunAsync(s, 81, filters)).GroupBy(x => T(x, "DayInsured")).ToDictionary(g => g.Key, g => g.Sum(x => N(x, "TotalS"))) : null;
            foreach (var x in await repository.RunAsync(s, 80, filters))
            {
                var n = supp?.GetValueOrDefault(T(x, "DayInsured")) ?? 0;
                rows.Add([Day(x), T(x, "TotalP"), supp is null ? "" : ((int)n).ToString(En)]);
                principals += (int)N(x, "TotalP");
                supps += (int)n;
            }
        }
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new MbmTrkResult(e.Headers, rows, principals, supps, 0, 0, 0);
    }
}
