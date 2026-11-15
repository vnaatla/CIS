using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Domain.Errors;
using FluentValidation;
using static Cis.Application.Reports.MbmTrkService;

namespace Cis.Application.Reports;

/// <summary>Claim / Bordx > Claim Tracking 2 search (FrmClaimRpt2 cmdSearch_Click / SearchRecords).</summary>
/// <param name="Basis">1 Actual (Option1), 2 Approved (Option2), 3 Cost Detail (Option3).</param>
/// <param name="Report">Option4 … Option20 (4 Claim … 20 Takeover); ignored for Cost Detail.</param>
/// <param name="Filters">The criteria by name (the cis.ClmTrk2_Run parameters without "@"); dates as dd/MM/yyyy, flags as true / false.</param>
public sealed record ClmTrk2Request(int Basis, int Report, Dictionary<string, string?> Filters);

/// <summary>Lengths.</summary>
public sealed class ClmTrk2RequestValidator : Common.RequestValidator<ClmTrk2Request>
{
    /// <summary>Defines the rules.</summary>
    public ClmTrk2RequestValidator()
    {
        RuleFor(r => r.Basis).InclusiveBetween(1, 3);
        RuleFor(r => r.Report).InclusiveBetween(0, 20);
        RuleFor(r => r.Filters).Must(f => f.Count <= 150 && f.All(x => x.Key.Length <= 40 && (x.Value ?? "").Length <= 300));
    }
}

/// <summary>A Claim Tracking 2 report choice.</summary>
public sealed record ClmTrk2Report(int Basis, int Report, string Name);

/// <summary>Data access.</summary>
public interface IClmTrk2Repository
{
    /// <summary>cis.ClmTrk2_Run: one case of CISClaimTracking2 with the typed filters.</summary>
    Task<List<Dictionary<string, object?>>> RunAsync(IDbSession session, int caseNo, IReadOnlyDictionary<string, object?> filters);
}

/// <summary>Claim / Bordx > Claim Tracking 2 (FrmClaimRpt2) and its Jupiter copy (FrmNewClaimRpt2).</summary>
public interface IClmTrk2Service
{
    /// <summary>The report choices (Actual / Approved x report, and Cost Detail).</summary>
    IReadOnlyList<ClmTrk2Report> Reports(string variant = "main");

    /// <summary>cmdSearch_Click.</summary>
    /// <param name="request">The search.</param>
    /// <param name="variant">main (FrmClaimRpt2) or jupiter (FrmNewClaimRpt2: agent / policy / patient compared exactly).</param>
    Task<ClmTrkResult> SearchAsync(ClmTrk2Request request, string variant = "main");
}

/// <inheritdoc />
public sealed class ClmTrk2Service(IDbSessions sessions, IClmTrk2Repository repository) : IClmTrk2Service
{
    /// <summary>A report: amount basis, option, routine, case, columns and the desktop headers.</summary>
    internal sealed record Entry(int Basis, int Report, string Name, string Routine, int C, List<Col> Cols, List<string> Headers);

    internal static readonly Dictionary<string, List<Entry>> Catalogues = new()
    {
        ["main"] = Load<List<Entry>>("ClmTracking2.catalogue.json"),
        ["jupiter"] = Load<List<Entry>>("ClmTracking2Jupiter.catalogue.json"),
    };

    private static readonly Dictionary<string, string> Params = Load<List<Param>>("ClmTracking2.params.json").ToDictionary(p => p.N, p => p.T, StringComparer.OrdinalIgnoreCase);
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;

    /// <summary>The 28 fee columns in the desktop's order (also the Cost Detail rows).</summary>
    internal static readonly string[] Fees =
    [
        "CLMGPFeeBasicMed", "CLMGPFeeMedInj", "CLMGPFeeLTMed", "CLMSPFeeBasicMed", "CLMSPFeeMedInj", "CLMSPFeeLTMed", "CLMXRay", "CLMUltraSound", "CLMECG", "CLMDressing",
        "CLMNebulizer", "CLMID", "CLMTSSTO", "CLMMinorSurgery", "CLMPOthers", "CLMBloodDrawing", "CLMBloodGlucose", "CLMUrine", "CLMPapSmear", "CLMLOthers", "CLMDTPPolio",
        "CLMDTPPolioHIB", "CLMHepatitisB", "CLMMMR", "CLMIOthers", "CLMDental", "CLMMaternity", "CLMPRVCheckup",
    ];

    private static readonly string[] CostLabels =
    [
        "101 - GP Fee & Basic Med", "102 - GP Fee & Med & Inj", "103 - GP Fee & LT Med", "104 - SP Fee & Basic Med", "105 - SP Fee & Med & Inj", "106 - SP Fee & LT Med",
        "201 - X-Ray", "202 - Ultra Sound", "203 - ECG", "204 - Dressing", "205 - Nebulizer", "206 - I&D/T&S", "207 - STO", "208 - Minor Surgery", "209 - Procedure Others",
        "301 - Blooding Drawing", "302 - Blood Glucose", "303 - Urine", "304 - Pap Smear", "305 - Lab Others", "401 - DTP & Polio", "402 - DTP & Polio & HIB",
        "403 - Hepatitis B", "404 - MMR", "405 - Immunization Others", "501 - Dental", "502 - Maternity", "503 - Preventive Checkup",
    ];

    /// <summary>Cost Detail columns: Actual (case 37), Approved (38), Pay to clinic (39), Pay to / from member (40).</summary>
    private static readonly int[] CostCases = [37, 38, 39, 40];

    /// <inheritdoc />
    public IReadOnlyList<ClmTrk2Report> Reports(string variant = "main") =>
        [.. Of(variant).Select(e => new ClmTrk2Report(e.Basis, e.Report, e.Name)), new ClmTrk2Report(3, 0, "Cost Detail")];

    private static List<Entry> Of(string variant) => Catalogues.TryGetValue(variant, out var c) ? c : throw new NotFoundException("Unknown tracking screen.");

    /// <summary>The typed filters (only the known names); the Jupiter copy compares agent / policy / patient exactly.</summary>
    internal static Dictionary<string, object?> Filters(Dictionary<string, string?> input, string variant)
    {
        var f = MbmTrkService.Filters(input, Params);
        f.Remove("exact");
        if (variant == "jupiter") f["exact"] = true;
        return f;
    }

    /// <inheritdoc />
    public async Task<ClmTrkResult> SearchAsync(ClmTrk2Request r, string variant = "main")
    {
        var catalogue = Of(variant);
        var filters = Filters(r.Filters, variant);
        await using var s = await sessions.OpenAsync();
        if (r.Basis == 3) return await CostDetailAsync(s, filters);
        var e = catalogue.FirstOrDefault(x => x.Basis == r.Basis && x.Report == r.Report) ?? throw new BusinessRuleException(NoReport);
        var rows = (await repository.RunAsync(s, e.C, filters)).Select(x => e.Cols.Select(c => c.K switch
        {
            "m0" => T(x, c.N!) is { Length: > 0 } ? MbmTrkService.N(x, c.N!).ToString("#,##0.00", En) : "0",
            // The row total: the 28 fee amounts added up.
            "sum" => Fees.Sum(k => MbmTrkService.N(x, k)).ToString("#,##0.00", En),
            "d" => x.GetValueOrDefault(c.N!) is DateTime d ? d.ToString("dd/MM/yyyy", En) : T(x, c.N!),
            "fe" => T(x, c.N!) is { Length: > 0 } v ? v : "Empty",
            "f" => T(x, c.N!),
            _ => "",
        }).ToList()).ToList();
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new ClmTrkResult(e.Headers, rows, rows.Count);
    }

    /// <summary>Cost Detail (CostDetailActual / Approved / CLN / MBM): one row per fee category with the four totals, and the total row.</summary>
    private async Task<ClmTrkResult> CostDetailAsync(IDbSession s, Dictionary<string, object?> filters)
    {
        var columns = new List<IReadOnlyDictionary<string, object?>?>();
        foreach (var c in CostCases) columns.Add((await repository.RunAsync(s, c, filters)).FirstOrDefault());
        var rows = CostLabels.Select((label, i) => new List<string> { label }.Concat(columns.Select(x => x is null ? "" : T(x, Fees[i]) is { Length: > 0 } v ? v : "0")).ToList()).ToList();
        rows.Add(["Total", .. columns.Select(x => x is null ? "" : Fees.Sum(k => MbmTrkService.N(x, k)).ToString("#,##0.00", En))]);
        return new ClmTrkResult(["Benefits", "Actual", "Approved", "Pay2Clinic", "Pay2/fr Mem"], rows, rows.Count);
    }
}
