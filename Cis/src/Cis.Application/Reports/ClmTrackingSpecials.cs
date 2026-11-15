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
        // ReportCLMGeneralListing (Monthly Reports > Claims Listing): department, case, clinic, five treatment reasons, GP / specialist amounts.
        ["ReportCLMGeneralListing"] = new(1, (x, _) =>
        {
            (string, string) Reason(string code, string desc) => Has(x, code) ? (F(x, code), F(x, desc)) : ("Empty", "Empty");
            var net = M(MbmTrkService.N(x, "CLMBordxAmt") - MbmTrkService.N(x, "ClmNonGPFee"));
            var specialist = F(x, "CLNType") == "S";
            var r1 = Reason("CASTreatReason1", "ReasonDesc");
            var r2 = Reason("CASTreatReason2", "ReasonDesc2");
            var r3 = Reason("CASTreatReason3", "ReasonDesc3");
            var r4 = Reason("CASTreatReason4", "ReasonDesc4");
            // Reason 5: codes 199 / 299 show the free-text diagnosis instead of the reason description.
            var r5 = !Has(x, "CASTreatReason5") ? ("Empty", "Empty")
                : F(x, "CASTreatReason5") is "199" or "299" ? (F(x, "CASTreatReason5"), F(x, "CASDiagnosis5")) : (F(x, "CASTreatReason5"), F(x, "ReasonDesc5"));
            return
            [
                Fe(x, "MBMDepartment"), Fe(x, "CASNumber"), F(x, "Bordxrefno"), D(x, "BordxDate"), F(x, "MBMNumber"), F(x, "coverid"), F(x, "PatientName"),
                F(x, "ClinicCode"), F(x, "CLNName"), F(x, "INVPayTo"), r1.Item1, r1.Item2, Has(x, "CLMTotalActual") ? M(MbmTrkService.N(x, "CLMTotalActual")) : "",
                specialist ? "0.00" : net, specialist ? net : "0.00", M0(x, "ClmNonGPFee"), D(x, "DateVisit"), Has(x, "MCDays") ? F(x, "MCDays") : "0",
                r2.Item1, r2.Item2, r3.Item1, r3.Item2, r4.Item1, r4.Item2, r5.Item1, r5.Item2,
            ];
        }),
        // ReportCLNSummary (Monthly Reports > Summary (by Clinics)): visits, MC times / days, bill amount and the per-visit figures.
        ["ReportCLNSummary"] = new(2, (x, _) =>
        {
            decimal visits = MbmTrkService.N(x, "NoofVisit"), times = MbmTrkService.N(x, "MCTimes"), days = MbmTrkService.N(x, "MCDays");
            return
            [
                Fe(x, "ClinicCode"), Fe(x, "CLNName"), F(x, "CLNType"), Has(x, "NoofVisit") ? F(x, "NoofVisit") : "0",
                times != 0 ? F(x, "MCTimes") : "0", days != 0 ? F(x, "MCDays") : "0", M0(x, "BillAmt"),
                Has(x, "BillAmt") && visits != 0 ? M(MbmTrkService.N(x, "BillAmt") / visits) : "0.00",
                times != 0 && visits != 0 ? R2(times / visits * 100) : "0",
                days != 0 ? (times != 0 ? M(days / times) : "0.00") : "0",
            ];
        }),
        // UtilizationByMBM (Utilization > Summary (By Member)): claims, MC, annual limit and the balance (FULL when unlimited).
        ["UtilizationByMBM"] = new(79, (x, _) =>
        {
            var limit = Has(x, "AnnLimit") ? F(x, "AnnLimit") : "FULL";
            var approved = Has(x, "ApprovedAmt") ? MbmTrkService.N(x, "ApprovedAmt") : 0;
            var balance = limit == "FULL" ? "FULL"
                : decimal.TryParse(limit, NumberStyles.Any, En, out var l) ? M(l - Math.Round(approved, 2, MidpointRounding.AwayFromZero)) : limit;
            return
            [
                Fe(x, "MBMNumber"), Fe(x, "MBMName"), Has(x, "TotalClaim") ? F(x, "TotalClaim") : "0", Has(x, "MCStatus") ? F(x, "MCStatus") : "0",
                Has(x, "MCDays") ? F(x, "MCDays") : "0", limit, M0(x, "ActualAmt"), M0(x, "ApprovedAmt"), balance,
            ];
        }),
    };

    /// <summary>The listings built from several rows per output row or several queries.</summary>
    private static readonly string[] Composite = ["ReportMBMSummaryOS", "ReportMBMSummaryOS1", "DailyVisitListing", "MonthyReportListing", "YearListing", "CLNByTypePanel", "PlanByPanel", "ReportMBMGeneralListing", "ReportDepartmentalBordx", "ReportIllnessSummary", "UtilizationByIllness"];

    /// <summary>Desktop texts of the admission reports.</summary>
    public const string SameMonth = "Must key in same Month", NoDov = "Date of Visit can't be empty", NoYear = "Please enter Year";

    /// <summary>Whether a catalogue entry is ported by hand here.</summary>
    internal static bool IsHandPorted(string routine) => Specials.ContainsKey(routine) || Composite.Contains(routine);

    private static decimal R2m(decimal v) => Math.Round(v, 2, MidpointRounding.AwayFromZero);

    /// <summary>Monthly Reports > Summary (by Illness) and Utilization > Summary (By Illness): per treatment reason (case 6) its description, bill and non-GP
    /// amounts, and the GP / SP (/ others) visits and amounts from case 11 for that reason and clinic type, with the averages. Case 11 has no reason column, so
    /// it runs per reason as the desktop. The desktop's helpers tested report types 15 / 21 (the reports are 16 / 22) and left these columns empty; they are
    /// filled here. Known desktop issue kept: the illness summary's % is of the running bill total (the rows above and this one).</summary>
    private async Task<List<List<string>>> IllnessRowsAsync(IDbSession s, bool utilization, Dictionary<string, object?> filters)
    {
        var reasons = (await repository.LookupAsync(s, "reasons", null)).Select(x => x.Split('|', 2)).GroupBy(x => x[0]).ToDictionary(g => g.Key, g => g.First()[1]);
        var userType = filters.GetValueOrDefault("clinicType") as string;
        async Task<(decimal Visits, decimal Amount)> Part(string code, string clinicType)
        {
            if (userType is not null && userType != clinicType) return (0, 0);
            var f = new Dictionary<string, object?>(filters, StringComparer.OrdinalIgnoreCase) { ["clinicType"] = clinicType };
            f.Remove("treatmentFrom"); f.Remove("treatmentTo"); f.Remove("treatmentBlank");
            if (code.Length == 0) f["treatmentBlank"] = true; else { f["treatmentFrom"] = code; f["treatmentTo"] = code; }
            var r = (await RunAsync(s, 11, f)).FirstOrDefault();
            return r is null ? (0, 0) : (MbmTrkService.N(r, "Total"), MbmTrkService.N(r, "TotalApproved") - MbmTrkService.N(r, "ClmNonGPFee"));
        }
        static string Num(decimal v) => v.ToString(CultureInfo.InvariantCulture);
        static string Avg(decimal amount, decimal visits, bool or) => (or ? amount > 0 || visits > 0 : amount > 0 && visits > 0) && visits != 0 ? R2(amount / visits) : "0";
        var rows = new List<List<string>>();
        decimal runningBill = 0;
        foreach (var x in await RunAsync(s, 6, filters))
        {
            var code = F(x, "CASTreatReason1");
            var desc = reasons.GetValueOrDefault(code) ?? "";
            var (gv, ga) = await Part(code, "1");
            var (sv, sa) = await Part(code, "0");
            if (!utilization)
            {
                runningBill += MbmTrkService.N(x, "CLMTotalActual");
                var bill = MbmTrkService.N(x, "CLMTotalActual");
                rows.Add([Fe(x, "CASTreatReason1"), desc, Num(gv), Num(sv), Num(gv + sv), M(ga), M(sa), M0(x, "ClmNonGPFee"),
                    Has(x, "CLMTotalActual") ? M(bill) : "0.00", Has(x, "CLMTotalActual") && runningBill != 0 ? R2(bill / runningBill * 100) : "0.00",
                    Avg(ga, gv, false), Avg(sa, sv, false), Avg(ga + sa, gv + sv, false)]);
            }
            else
            {
                var (ov, oa) = await Part(code, "2");
                rows.Add([Fe(x, "CASTreatReason1"), desc, Num(gv), Num(sv), Num(ov), Num(gv + sv + ov), M(ga), M(sa), M(oa), M0(x, "ClmNonGPFee"),
                    M0(x, "CLMTotalActual"), Num(ga + sa + oa), Avg(ga, gv, true), Avg(sa, sv, true), Avg(oa, ov, true), Avg(ga + sa + oa, gv + sv + ov, true)]);
            }
        }
        return rows;
    }

    /// <summary>Utilisation % of the annual limit (with a negative top-up added back), as the O/S Utilized reports.</summary>
    private static decimal? Util(decimal approved, decimal topUp, IReadOnlyDictionary<string, object?> x)
    {
        var limit = MbmTrkService.N(x, "AnnLimit");
        var used = topUp < 0 ? -topUp + approved : approved;
        return limit > 0 ? Math.Round(used / limit * 100, 2, MidpointRounding.ToEven) : 0;
    }

    /// <summary>ReportMBMSummaryOS (combined, case 5): one row per member (consecutive rows summed); ReportMBMSummaryOS1 (separate, case 44): one row per row,
    /// with cover id and patient. Known desktop issue kept in OS1: its "FULL" test reads the MCO fee column, so the utilisation is always computed.</summary>
    private async Task<List<List<string>>> CompositeRowsAsync(IDbSession s, Entry e, Dictionary<string, object?> filters)
    {
        var separate = e.Routine == "ReportMBMSummaryOS1";
        var rows = new List<List<string>>();
        List<string>? row = null;
        string member = "\u0001";
        decimal claims = 0, times = 0, days = 0, approved = 0, topUp = 0;
        foreach (var x in await RunAsync(s, separate ? 44 : 5, filters))
        {
            var y = F(x, "AnnLmtIND") == "Y";
            if (separate || F(x, "MBMNumber") != member)
            {
                member = F(x, "MBMNumber");
                topUp = y && Has(x, "TopUpAmt") ? R2m(MbmTrkService.N(x, "TopUpAmt")) : 0;
                row = separate
                    ? [Fe(x, "MBMDepartment"), F(x, "MBMNumber"), F(x, "coverid"), F(x, "PatientName"), "", "", "", y ? M0(x, "AnnLimit") : "FULL",
                        y ? M0(x, "TopUpAmt") : "FULL", "", M0(x, "CLMMCOProcFee"), y ? M(MbmTrkService.N(x, "MBMCLNBalLmt")) : "FULL", "FULL"]
                    : [Fe(x, "MBMDepartment"), F(x, "MBMNumber"), F(x, "MBMName"), "", "", "", y ? M0(x, "AnnLimit") : "FULL",
                        y && Has(x, "TopUpAmt") ? M(MbmTrkService.N(x, "TopUpAmt")) : "0.00", "", M0(x, "CLMMCOProcFee"),
                        y ? M(MbmTrkService.N(x, "MBMCLNBalLmt")) : "FULL", y ? "" : "FULL"];
                rows.Add(row);
                claims = times = days = approved = 0;
            }
            var o = separate ? 1 : 0;
            claims += MbmTrkService.N(x, "TotalClaim");
            times += MbmTrkService.N(x, "MCTimes");
            days += MbmTrkService.N(x, "MCDays");
            approved += R2m(MbmTrkService.N(x, "ApprovedAmt"));
            row![3 + o] = Has(x, "TotalClaim") ? ((int)claims).ToString(En) : "0";
            row[4 + o] = Has(x, "MCTimes") ? ((int)times).ToString(En) : "0";
            row[5 + o] = Has(x, "MCDays") ? ((int)days).ToString(En) : "0";
            row[8 + o] = Has(x, "ApprovedAmt") ? M(approved) : "0.00";
            var util = 11 + o;
            if (!separate && row[6] == "FULL") row[util] = "FULL";
            else if (approved > 0) row[util] = separate && !Has(x, "AnnLimit") ? "0" : M(Util(approved, topUp, x)!.Value);
            else row[util] = "0.00";
        }
        return rows;
    }

    private async Task<List<List<string>>> SpecialRowsAsync(IDbSession s, Entry e, Dictionary<string, object?> filters)
    {
        if (e.Routine == "DailyVisitListing")
            // Admission > Daily Report (case 34): the visits numbered 1, 2, ... (the desktop's counter was never reset, so it kept counting across searches).
            return (await RunAsync(s, 34, filters)).Select((x, i) => new List<string>
            {
                (i + 1).ToString(En), F(x, "CASNumber"), F(x, "MBMNumber"), F(x, "MBMPolicyNo"), F(x, "PatientName"), D(x, "DateVisit"), F(x, "ClinicCode"),
                F(x, "CLNName"), F(x, "CASTreatReason1"), F(x, "MCDays"),
            }).ToList();
        if (e.Routine == "MonthyReportListing")
        {
            // Admission > Monthly Report: DOV from / to in the same month (the desktop compared the mm/yyyy of the two boxes first).
            var from = filters.GetValueOrDefault("dovFrom") as DateTime?;
            var to = filters.GetValueOrDefault("dovTo") as DateTime?;
            if ((from is null) != (to is null) || (from is not null && (from.Value.Month != to!.Value.Month || from.Value.Year != to.Value.Year)))
                throw new Domain.Errors.BusinessRuleException(SameMonth);
            if (from is null) throw new Domain.Errors.BusinessRuleException(NoDov);
            // Per payor (case 35) its visit days in order (case 96): "<date>-<weekday>" and the visits, up to 31 days. The desktop did not clear the previous
            // payor's days, so a payor with fewer days showed the remaining days of the one before; here each payor shows its own days.
            var days = (await RunAsync(s, 96, filters)).GroupBy(x => F(x, "PAYCode")).ToDictionary(g => g.Key, g => g.Take(31).ToList());
            return (await RunAsync(s, 35, filters)).Where(x => Has(x, "PAYCode")).Select(x =>
            {
                var row = new List<string> { F(x, "PAYCode") };
                foreach (var d in days.GetValueOrDefault(F(x, "PAYCode")) ?? [])
                {
                    var dov = d.GetValueOrDefault("DateVisit") as DateTime?;
                    row.Add(dov is { } v ? v.ToString("dd/MM/yyyy", En) + "-" + v.DayOfWeek : "");
                    row.Add(F(d, "Total"));
                }
                return row;
            }).ToList();
        }
        if (e.Routine == "YearListing")
        {
            // Admission > YTD Report: the year is required; per payor (case 36) the visits of each month 1-12 (case 97, the first row of a month).
            if (!filters.ContainsKey("year")) throw new Domain.Errors.BusinessRuleException(NoYear);
            var months = (await RunAsync(s, 97, filters)).GroupBy(x => F(x, "PAYCode")).ToDictionary(g => g.Key, g => g.ToList());
            return (await RunAsync(s, 36, filters)).Where(x => Has(x, "PAYCode")).Select(x =>
            {
                var mine = months.GetValueOrDefault(F(x, "PAYCode")) ?? [];
                var row = new List<string> { F(x, "PAYCode") };
                for (var m = 1; m <= 12; m++) row.Add(mine.FirstOrDefault(r => F(r, "Month") == m.ToString(En)) is { } r ? F(r, "TotalCases") : "");
                return row;
            }).ToList();
        }
        if (e.Routine is "CLNByTypePanel" or "PlanByPanel")
        {
            // Clinic type (case 24) / plan code (case 82) rows with their panel (cases 31 / 12, CLNPanel C) and non-panel (32 / 33) claims and amounts.
            // The desktop ran the two queries per row with the row's type / plan added; each runs once here, joined on it. The user's own panel criterion was
            // appended too: panel C leaves no non-panel rows, panel Y no panel rows.
            var byType = e.Routine == "CLNByTypePanel";
            var key = byType ? "CLNType" : "PLNCode";
            var userPanel = filters.GetValueOrDefault("panel") as string;
            async Task<Dictionary<string, IReadOnlyDictionary<string, object?>>> Part(int c, bool panel)
            {
                if (userPanel is not null && userPanel == "C" != panel) return [];
                var f = new Dictionary<string, object?>(filters, StringComparer.OrdinalIgnoreCase);
                if (panel) f["panel"] = "C"; else f["nonPanel"] = true;
                return (await RunAsync(s, c, f)).GroupBy(x => F(x, key)).ToDictionary(g => g.Key, g => (IReadOnlyDictionary<string, object?>)g.First());
            }
            var panelRows = await Part(byType ? 31 : 12, true);
            var nonPanelRows = await Part(byType ? 32 : 33, false);
            return (await RunAsync(s, byType ? 24 : 82, filters)).Select(x =>
            {
                var p = panelRows.GetValueOrDefault(F(x, key));
                var n = nonPanelRows.GetValueOrDefault(F(x, key));
                decimal Of(IReadOnlyDictionary<string, object?>? r, string k) => r is null ? 0 : MbmTrkService.N(r, k);
                if (!byType) return new List<string> { Fe(x, "PLNCode"), M0(x, "TotalApprovedAmt"), M(Of(p, "TotalPanelAmt")), M(Of(n, "TotalNonPanelAmt")) };
                decimal pc = Of(p, "TotalClaimP"), pa = Of(p, "TotalPanelAmt"), nc = Of(n, "TotalClaimNP"), na = Of(n, "TotalNonPanelAmt");
                return new List<string> { Fe(x, "CLNType"), ((int)pc).ToString(En), M(pa), ((int)nc).ToString(En), M(na), ((int)(pc + nc)).ToString(En), M(pa + na) };
            }).ToList();
        }
        if (e.Routine is "ReportMBMGeneralListing" or "ReportDepartmentalBordx")
        {
            // Monthly Reports > Utilization (By Member) (case 4 + case 108) / Departmental Bordx (case 3 + case 107): the GP and SP amounts (bill less non-GP fee)
            // by clinic type. The desktop's helper tested report type 15 (Monthly Reports is 16), so these columns stayed empty there; they are filled here.
            var member = e.Routine == "ReportMBMGeneralListing";
            string[] keys = member ? ["MBMDepartment", "MBMNumber", "CoverID"] : ["MBMDepartment", "BordxRefNo"];
            string K(IReadOnlyDictionary<string, object?> x) => string.Join("\u0001", keys.Select(k => F(x, k)));
            var gp = new Dictionary<string, string>();
            var spAmt = new Dictionary<string, string>();
            foreach (var x in await RunAsync(s, member ? 108 : 107, filters))
                (F(x, "CLNType") == "S" ? spAmt : gp)[K(x)] = M(MbmTrkService.N(x, "CLMTotalActual") - MbmTrkService.N(x, "ClmNonGPFee"));
            return (await RunAsync(s, member ? 4 : 3, filters)).Select(x => member
                ? new List<string> { Fe(x, "MBMDepartment"), F(x, "MBMNumber"), F(x, "coverid"), F(x, "PatientName"), Has(x, "TotalVisit") ? F(x, "TotalVisit") : "0",
                    Has(x, "MCTimes") ? F(x, "MCTimes") : "0", Has(x, "MCDays") ? F(x, "MCDays") : "0", M0(x, "CLMTotalActual"), gp.GetValueOrDefault(K(x)) ?? "",
                    spAmt.GetValueOrDefault(K(x)) ?? "" }
                : new List<string> { Fe(x, "MBMDepartment"), F(x, "Bordxrefno"), D(x, "BordxDate"), M0(x, "CLMTotalActual"), gp.GetValueOrDefault(K(x)) ?? "",
                    spAmt.GetValueOrDefault(K(x)) ?? "", M0(x, "ClmNonGPFee") }).ToList();
        }
        if (e.Routine is "ReportIllnessSummary" or "UtilizationByIllness")
            return await IllnessRowsAsync(s, e.Routine == "UtilizationByIllness", filters);
        if (Composite.Contains(e.Routine)) return await CompositeRowsAsync(s, e, filters);
        var sp = Specials[e.Routine];
        var reasons = sp.Reasons
            ? (await repository.LookupAsync(s, "reasons", null)).Select(x => x.Split('|', 2)).GroupBy(x => x[0]).ToDictionary(g => g.Key, g => g.First()[1])
            : new Dictionary<string, string>();
        return (await RunAsync(s, sp.Case, filters)).Select(x => sp.Map(x, reasons)).ToList();
    }
}
