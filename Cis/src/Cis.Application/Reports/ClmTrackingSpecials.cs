using System.Globalization;
using Cis.Application.Abstractions;
using static Cis.Application.Reports.MbmTrkService;

namespace Cis.Application.Reports;

/// <summary>The claim tracking listings whose columns follow row conditions (ported by hand from frmClaimRpt).</summary>
public sealed partial class ClmTrkService
{
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;
    private delegate List<string> RowMap(IReadOnlyDictionary<string, object?> x, IReadOnlyDictionary<string, string> reasons);

    private sealed record Special(int Case, RowMap Map, bool Reasons = false);

    private static string F(IReadOnlyDictionary<string, object?> x, string k) => T(x, k);
    private static string Fe(IReadOnlyDictionary<string, object?> x, string k) => T(x, k) is { Length: > 0 } v ? v : "Empty";
    private static string D(IReadOnlyDictionary<string, object?> x, string k) => x.GetValueOrDefault(k) is DateTime d ? d.ToString("dd/MM/yyyy", En) : T(x, k);
    private static bool Has(IReadOnlyDictionary<string, object?> x, string k) => T(x, k).Length > 0;
    private static string M(decimal v) => v.ToString("#,##0.00", En);
    /// <summary>Format(v, "#,##0.00;(#,##0.00)"), "0.00" when blank.</summary>
    private static string Mp0(IReadOnlyDictionary<string, object?> x, string k) => Has(x, k) ? MbmTrkService.N(x, k).ToString("#,##0.00;(#,##0.00)", En) : "0.00";
    private static string M0(IReadOnlyDictionary<string, object?> x, string k) => Has(x, k) ? M(MbmTrkService.N(x, k)) : "0.00";
    private static string R2(decimal v) => Math.Round(v, 2, MidpointRounding.ToEven).ToString(En);

    private static readonly Dictionary<string, Special> Specials = new()
    {
        // AgentListing: the agent's claims, with the first treatment reason's description.
        ["AgentListing"] = new(17, (x, r) =>
        [
            Fe(x, "PAYCode"), Fe(x, "MBMAgentCode"), F(x, "CASNumber"), F(x, "MBMNumber"), F(x, "coverid"), F(x, "PatientName"), "",
            Mp0(x, "CLMTotalActual"), Mp0(x, "CLMTotalApproved"), Mp0(x, "CLMRecoveryPaidByM"), F(x, "CASTreatReason1"),
            r.GetValueOrDefault(F(x, "CASTreatReason1")) ?? "", D(x, "DateVisit"), F(x, "MCDays"), F(x, "INSCode"), F(x, "PLNCode"), F(x, "CASStatus"),
        ], true),
        // CLNSumListing2: per clinic, averages per visit and the MC rates.
        ["CLNSumListing2"] = new(27, (x, _) =>
        {
            var total = MbmTrkService.N(x, "total");
            var mcTimes = MbmTrkService.N(x, "MCTimes");
            var row = new List<string> { Fe(x, "ClinicCode"), Fe(x, "CLNName"), Fe(x, "CLNType"), Fe(x, "CLNLocation"), F(x, "total"),
                Has(x, "MCTimes") ? F(x, "MCTimes") : "0", F(x, "MCDays"), M0(x, "TotalActual"), M0(x, "TotalApproved") };
            row.AddRange(total != 0 ? [M(Math.Round(MbmTrkService.N(x, "TotalActual") / total, 2, MidpointRounding.ToEven)), M(Math.Round(MbmTrkService.N(x, "TotalApproved") / total, 2, MidpointRounding.ToEven))] : ["0.00", "0.00"]);
            row.AddRange(total != 0 && mcTimes != 0 ? [R2(mcTimes / total * 100), R2(MbmTrkService.N(x, "MCDays") / mcTimes)] : ["0", "0"]);
            return row;
        }),
        // ResultListing (General Listing without remarks): the claims with member, patient, clinic, amounts and audit columns.
        ["ResultListing"] = new(93, (x, _) =>
        {
            var reason = F(x, "CASTreatReason1");
            var (code, desc) = reason is "016" or "032" ? (reason, F(x, "CASDiagnosis5")) : reason.Length > 0 ? (reason, F(x, "ReasonDesc")) : ("Empty", "Empty");
            var net = M(MbmTrkService.N(x, "CLMTotalActual") - MbmTrkService.N(x, "ClmNonGPFee"));
            var specialist = F(x, "CLNType") == "S";
            var payable = Has(x, "CLMPayableByMedix2M") && MbmTrkService.N(x, "CLMPayableByMedix2M") < 0 ? MbmTrkService.N(x, "CLMPayableByMedix2M").ToString("#,##0.00;(#,##0.00)", En) : "0.00";
            return
            [
                Fe(x, "PAYCode"), F(x, "CASNumber"), F(x, "Bordxrefno"), D(x, "BordxDate"), F(x, "MBMNumber"), F(x, "MBMPolicyNo"), F(x, "coverid"), F(x, "MBMName"),
                F(x, "PatientName"), F(x, "MBMIcBcPp"), F(x, "MBMSex"), F(x, "AgeCode"), D(x, "MBMDOB"), F(x, "IcNo"), F(x, "Sex"), F(x, "AgeBand"), D(x, "DOB"),
                F(x, "INSCode"), F(x, "PLNCode"), F(x, "ClinicCode"), F(x, "CLNName"), code, desc, Mp0(x, "CLMTotalActual"), Mp0(x, "CLMTotalApproved"), payable,
                Mp0(x, "MBMCLNPremium"), F(x, "CASStatus"), F(x, "Claimability"), F(x, "MBMStatus"), F(x, "MBMDataStatus"), D(x, "DateVisit"), D(x, "RegisterDate"),
                F(x, "MCDays"), D(x, "ClaimOpenDate"), D(x, "MBMPayorEffDate"), D(x, "MBMPayorEXPDate"), F(x, "MBMDepartment"), F(x, "INVPayTo"),
                specialist ? "0.00" : net, specialist ? net : "0.00", Mp0(x, "ClmNonGPFee"), F(x, "InvoiceNo"), F(x, "MBMGroupCompany"), F(x, "ReasonDesc11"),
                F(x, "ReasonDesc22"), F(x, "ReasonDesc33"), F(x, "ReasonDesc44"), F(x, "ReasonDesc55"), F(x, "RegisterBy"), F(x, "LastUpdateBy"), F(x, "BTBGBordxRefNo"),
                F(x, "CLMBTBGAmt"), F(x, "CLMBordxAmt"),
            ];
        }),
    };

    /// <summary>Whether a catalogue entry is ported by hand here.</summary>
    internal static bool IsHandPorted(string routine) => Specials.ContainsKey(routine);

    private async Task<List<List<string>>> SpecialRowsAsync(IDbSession s, Entry e, Dictionary<string, object?> filters)
    {
        var sp = Specials[e.Routine];
        var reasons = sp.Reasons
            ? (await repository.LookupAsync(s, "reasons", null)).Select(x => x.Split('|', 2)).GroupBy(x => x[0]).ToDictionary(g => g.Key, g => g.First()[1])
            : new Dictionary<string, string>();
        return (await repository.RunAsync(s, sp.Case, filters)).Select(x => sp.Map(x, reasons)).ToList();
    }
}
