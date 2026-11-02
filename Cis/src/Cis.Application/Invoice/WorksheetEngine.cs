namespace Cis.Application.Invoice;

/// <summary>Which invoice screen opened the worksheet (the desktop's INVCLNForm).</summary>
public enum WorksheetCaller
{
    /// <summary>Clinic invoice (frmInvoice).</summary>
    ClinicInvoice = 1,
    /// <summary>Member invoice (frmInvoiceMBM).</summary>
    MemberInvoice = 2,
    /// <summary>Adjustment (frmInvoiceAdj).</summary>
    Adjustment = 3,
    /// <summary>Adjustment after bordereaux (frmInvoiceAdjAftBordx).</summary>
    AdjustmentAfterBordx = 4,
    /// <summary>Adjustment after posting (frmInvoiceAdjAfterPost).</summary>
    AdjustmentAfterPost = 5,
}

/// <summary>The invoice row values the worksheet reads (Form_Load / showInvoice). Amounts that are not numbers in the row are null.</summary>
public sealed record WorksheetInput
{
    /// <summary>Calling screen.</summary>
    public WorksheetCaller Caller { get; init; } = WorksheetCaller.ClinicInvoice;
    /// <summary>Row 80: C cashless / M reimbursement.</summary>
    public string PayTo { get; init; } = "C";
    /// <summary>Row 12.</summary>
    public string ClinicCode { get; init; } = "";
    /// <summary>Row 72: the clinic benefit card (STDClinicBNFSPE type).</summary>
    public string ClinicBenefitCard { get; init; } = "";
    /// <summary>Row 67.</summary>
    public string PlanCode { get; init; } = "";
    /// <summary>Row 68.</summary>
    public string InsCode { get; init; } = "";
    /// <summary>Row 69.</summary>
    public string HltCode { get; init; } = "";
    /// <summary>I1..I31: the item amounts in grid order (row 15–20, 24–42, 21–23, 101 or 95, 474 or 473, 475).</summary>
    public IReadOnlyList<decimal?> Items { get; init; } = [];
    /// <summary>Row 114.</summary>
    public decimal NonGpAmount { get; init; }
    /// <summary>Row 70 (also shown for the clinic / member invoice) — the annual balance.</summary>
    public string Row70BalanceLimit { get; init; } = "";
    /// <summary>The "Balance Limit" shown: row 70 (invoices) or row 164 (adjustments).</summary>
    public string BalanceLimit { get; init; } = "";
    /// <summary>Row 71 ("FULL" or an amount).</summary>
    public string VisitLimit { get; init; } = "";
    /// <summary>Row 74 not "N".</summary>
    public bool AnnualIndicator { get; init; }
    /// <summary>Row 73 not "N".</summary>
    public bool VisitIndicator { get; init; }
    /// <summary>Row 43.</summary>
    public decimal CoPayCollected { get; init; }
    /// <summary>Row 75.</summary>
    public decimal CoPay { get; init; }
    /// <summary>Row 113.</summary>
    public decimal AdjustedAnnualLimit { get; init; }
    /// <summary>Row 91 &gt; 0.</summary>
    public bool GpStatus { get; init; }
    /// <summary>Row 469 = "Y".</summary>
    public bool McoProcStatus { get; init; }
    /// <summary>Adjustment screens: the stored contracted rates (row 195–225).</summary>
    public IReadOnlyList<decimal?>? StoredRates { get; init; }
    /// <summary>Adjustment screens: the stored benefit limits (row 297–327).</summary>
    public IReadOnlyList<decimal?>? StoredLimits { get; init; }
    /// <summary>Member invoice row 476 / adjustment (pay to M) row 475: an approved amount that overrides the worksheet when &gt; 0.</summary>
    public string? ApprovedOverride { get; init; }
}

/// <summary>A benefit of the plan (BNFNew).</summary>
/// <param name="Code">BNFCode.</param>
/// <param name="Percentage">BNFPercentage.</param>
/// <param name="RmPerVisit">BNFRMPerVisit.</param>
/// <param name="Calculation">BNFCalculation 1 min, 2 max, 3 actual, 4 RM per visit.</param>
public record WorksheetBenefit(string Code, decimal? Percentage, decimal? RmPerVisit, string? Calculation);

/// <summary>Database values the worksheet looks up.</summary>
public sealed record WorksheetLookups
{
    /// <summary>BNFCategoryNew (code, description) in code order — the 31 grid rows.</summary>
    public IReadOnlyList<(string Code, string Description)> Categories { get; init; } = [];
    /// <summary>CLN.CLNBNFType (1 actual billing, 2 standard rates …).</summary>
    public string ClinicBenefitType { get; init; } = "";
    /// <summary>CLN.CLNRedStatus (1 fixed amount, 2 percentage).</summary>
    public string ClinicReductionStatus { get; init; } = "";
    /// <summary>CLN.CLNMCOProcfeePerc.</summary>
    public decimal McoProcFeePercent { get; init; }
    /// <summary>STDClinicBNF rates by code.</summary>
    public IReadOnlyDictionary<string, decimal?> StandardRates { get; init; } = new Dictionary<string, decimal?>();
    /// <summary>STDClinicBNFSPE rates of the clinic's card by code.</summary>
    public IReadOnlyDictionary<string, decimal?> SpecialRates { get; init; } = new Dictionary<string, decimal?>();
    /// <summary>PLN.BNFStatus = "Y".</summary>
    public bool PlanBenefitStatus { get; init; }
    /// <summary>BNFNew rows of the plan / ins / hlt.</summary>
    public IReadOnlyList<WorksheetBenefit> Benefits { get; init; } = [];
    /// <summary>CLNRedAmt first row: amount and percentage.</summary>
    public decimal ReductionAmount { get; init; }
    /// <summary>CLNRedAmt first row percentage.</summary>
    public decimal ReductionPercent { get; init; }
}

/// <summary>A benefit row of the grid; null = shown as "-" (or empty).</summary>
/// <param name="Code">BNF code.</param>
/// <param name="Benefit">Description.</param>
/// <param name="Actual">Actual.</param>
/// <param name="ContractedRate">Contracted Rate.</param>
/// <param name="PayableToClinic">Payable to CLN.</param>
/// <param name="NotPayableToClinic">Not Payable to CLN.</param>
/// <param name="BenefitLimit">Benefit Lmt.</param>
/// <param name="Approved">Approved Benefit.</param>
/// <param name="Excess">Excess.</param>
public record WorksheetRow(string Code, string Benefit, decimal? Actual, decimal? ContractedRate, decimal? PayableToClinic, decimal? NotPayableToClinic,
    decimal? BenefitLimit, decimal? Approved, decimal? Excess);

/// <summary>The worksheet: grid rows, the total lines and the values written back to the invoice row (SaveCLM* / SaveClaim).</summary>
/// <param name="Rows">31 benefit rows.</param>
/// <param name="TotalActual">Total: actual.</param>
/// <param name="TotalPayableToClinic">Total: payable to clinic.</param>
/// <param name="TotalNotPayableToClinic">Total: not payable to clinic.</param>
/// <param name="TotalApproved">Total: approved.</param>
/// <param name="TotalExcess">Total: excess.</param>
/// <param name="CoPay">Less Co-Payment (shown).</param>
/// <param name="CoPayCollected">Less Co-Payment collected.</param>
/// <param name="VisitLimit">Visit Limit (text, e.g. FULL).</param>
/// <param name="BalanceLimit">Balance Limit (text).</param>
/// <param name="McoProcFee">MCO Proc. Fee.</param>
/// <param name="FinalPayableToClinic">Final Approved: payable to clinic.</param>
/// <param name="FinalApproved">Final Approved: approved (null = "-").</param>
/// <param name="FinalExcess">Final Approved: excess (member).</param>
/// <param name="FinalBalanceLimit">FinalBalLmt.</param>
/// <param name="RowValues">Invoice row SubItems to set (index → value).</param>
public record WorksheetResult(List<WorksheetRow> Rows, decimal TotalActual, decimal TotalPayableToClinic, decimal TotalNotPayableToClinic,
    decimal TotalApproved, decimal TotalExcess, decimal CoPay, decimal CoPayCollected, string VisitLimit, string BalanceLimit, decimal McoProcFee,
    decimal FinalPayableToClinic, decimal? FinalApproved, decimal FinalExcess, decimal FinalBalanceLimit, IReadOnlyDictionary<int, object?> RowValues);

/// <summary>frmWorksheet's calculation (Form_Load → showInvoice, display_benefit, display_CLNbenefit, getDatatoWS, getFinalTotal, Save*).</summary>
public static class WorksheetEngine
{
    /// <summary>The standard-rate codes per grid row used by display_CLNbenefit (101–106 from the card's special rates, the rest standard).</summary>
    private static readonly string[] SpecialCodes = ["101", "102", "103", "104", "105", "106"];

    /// <summary>Calculates the worksheet. <paramref name="today"/> and <paramref name="userCode"/> fill the claim audit fields (SaveClaim).</summary>
    public static WorksheetResult Calculate(WorksheetInput input, WorksheetLookups db, DateTime today, string userCode)
    {
        var n = db.Categories.Count;
        decimal I(int i) => i < input.Items.Count ? input.Items[i] ?? 0 : 0;
        var cashless = input.PayTo == "C";
        // CBInd is only read for cashless (display_CLNbenefit); otherwise it stays empty.
        var cbInd = cashless ? db.ClinicBenefitType.Trim() : "";
        var rates = ContractedRates(input, db, cbInd, n, I);
        var limits = BenefitLimits(input, db, n, I);

        var rows = new List<WorksheetRow>(n);
        for (var i = 0; i < n; i++)
        {
            var a = I(i);
            var c = I(i);
            // getDatatoWS: L = I, so the approved benefit is the actual amount (GetMin(A, L)).
            var approved = Min(a, c);
            var excess = DueToFromMember(input.PayTo, a, approved ?? 0);
            decimal? contracted = cbInd == "1" ? c : rates[i];
            decimal? payable = cbInd == "1" ? c : cbInd is "2" or "3" ? Min(c, I(i)) : null;
            decimal? notPayable = cbInd != "1" ? DueToFromClinic(c, rates[i] ?? 0) : null;
            rows.Add(new WorksheetRow(db.Categories[i].Code, db.Categories[i].Description, a == 0 ? null : a, contracted, payable, notPayable, limits[i],
                approved, excess));
        }

        var totalActual = rows.Sum(r => r.Actual ?? 0);
        var totalApprove = rows.Sum(r => r.Approved ?? 0);
        var totalCln = rows.Sum(r => r.PayableToClinic ?? 0);
        var totalClnExc = rows.Sum(r => r.NotPayableToClinic ?? 0);
        var totalMem = rows.Sum(r => r.Excess ?? 0);

        // getFinalTotal.
        decimal totalFCln = 0, mcoFee = 0;
        if (cashless)
        {
            var red = db.ClinicReductionStatus.Trim() switch
            {
                "1" => db.ReductionAmount,
                "2" => db.ReductionPercent * totalCln / 100,
                _ => 0m,
            };
            // IIf(totalCLN - COPayColl - CLNRedAmt, totalCLN - COPayColl, 0) - CLNRedAmt
            totalFCln = (totalCln - input.CoPayCollected - red != 0 ? totalCln - input.CoPayCollected : 0) - red;
            if (input.McoProcStatus && input.GpStatus)
            {
                mcoFee = db.McoProcFeePercent * totalFCln / 100;
                totalFCln -= mcoFee;
            }
        }
        var finalBal = !input.AnnualIndicator ? totalApprove : (Num(input.Row70BalanceLimit) is { } b ? b + input.AdjustedAnnualLimit : 0);
        var finalVisit = !input.VisitIndicator || input.VisitLimit.Trim().Equals("FULL", StringComparison.OrdinalIgnoreCase)
            ? totalApprove : Num(input.VisitLimit) ?? 0;
        decimal? finalApproved = finalVisit > finalBal
            ? totalApprove > finalBal ? finalBal - input.CoPayCollected : totalApprove != 0 ? totalApprove - input.CoPayCollected : null
            : totalApprove - input.CoPayCollected - input.CoPay;
        var totalFMem = totalMem - input.CoPayCollected;

        return new WorksheetResult(rows, totalActual, totalCln, totalClnExc, totalApprove, totalMem, input.CoPay, input.CoPayCollected,
            input.VisitLimit, input.BalanceLimit, mcoFee, totalFCln, finalApproved, totalFMem, finalBal,
            RowValues(input, rows, totalActual, totalFCln, finalApproved ?? 0, totalFMem, finalBal, mcoFee, today, userCode));
    }

    /// <summary>display_CLNbenefit's contracted rates (CB1..CB31).</summary>
    private static decimal?[] ContractedRates(WorksheetInput input, WorksheetLookups db, string cbInd, int n, Func<int, decimal> item)
    {
        var rates = new decimal?[n];
        if (input.PayTo != "C") return rates;
        if (input.Caller is WorksheetCaller.Adjustment or WorksheetCaller.AdjustmentAfterBordx or WorksheetCaller.AdjustmentAfterPost)
        {
            for (var i = 0; i < n; i++) rates[i] = input.StoredRates is { } s && i < s.Count ? s[i] ?? 0 : 0;
            return rates;
        }
        if (input.Caller != WorksheetCaller.ClinicInvoice) return rates;
        for (var i = 0; i < n; i++)
        {
            var code = db.Categories[i].Code;
            rates[i] = cbInd switch
            {
                "2" => SpecialCodes.Contains(code) ? db.SpecialRates.GetValueOrDefault(code) ?? 0 : db.StandardRates.TryGetValue(code, out var r) ? r ?? 0 : 0,
                "1" => item(i),
                _ => 0,
            };
        }
        return rates;
    }

    /// <summary>display_benefit's benefit limits (grid column 6).</summary>
    private static decimal?[] BenefitLimits(WorksheetInput input, WorksheetLookups db, int n, Func<int, decimal> item)
    {
        var limits = new decimal?[n];
        if (input.Caller is WorksheetCaller.ClinicInvoice or WorksheetCaller.MemberInvoice)
        {
            if (!db.PlanBenefitStatus)
            {
                for (var i = 0; i < n; i++) limits[i] = item(i);
                return limits;
            }
            foreach (var b in db.Benefits)
            {
                var i = db.Categories.ToList().FindIndex(c => c.Code == b.Code);
                if (i < 0) continue;
                var inv = (b.Percentage ?? 0) / 100 * item(i);
                var lmt = b.RmPerVisit ?? 0;
                limits[i] = b.Calculation?.Trim() switch
                {
                    "1" => Min(inv, lmt),
                    "2" => Max(inv, lmt),
                    "3" => item(i),
                    "4" => lmt,
                    _ => limits[i],
                };
            }
            return limits;
        }
        for (var i = 0; i < n; i++) limits[i] = input.StoredLimits is { } s && i < s.Count ? s[i] : null;
        return limits;
    }

    /// <summary>SaveCLMActual / ContRate / PayCLN / NotPayCLN / MBMBnfLmt / MBMApp / MBMExc and SaveClaim. The desktop copies the Actual column
    /// into the contracted-rate, payable, benefit-limit and approved slots (reproduced).</summary>
    private static Dictionary<int, object?> RowValues(WorksheetInput input, List<WorksheetRow> rows, decimal totalActual, decimal totalFCln,
        decimal finalApproved, decimal totalFMem, decimal finalBal, decimal mcoFee, DateTime today, string userCode)
    {
        var v = new Dictionary<int, object?>();
        for (var i = 0; i < rows.Count; i++)
        {
            var actual = rows[i].Actual ?? 0;
            v[168 + i] = actual;
            v[211 + i] = actual;
            v[426 + i] = actual;
            v[297 + i] = actual;
            v[254 + i] = actual;
            v[383 + i] = rows[i].NotPayableToClinic ?? 0;
            v[340 + i] = rows[i].Excess ?? 0;
        }
        var user = userCode.Trim().ToUpperInvariant();
        v[150] = today; v[151] = user; v[152] = user; v[153] = today; v[154] = 0m; v[155] = "C"; v[156] = user; v[157] = today;
        v[158] = totalFCln;
        // TotalFinalCLNExc is never set by the desktop.
        v[159] = 0m;
        v[160] = totalActual + input.NonGpAmount;
        var overrideApplies = input.Caller == WorksheetCaller.MemberInvoice
            || (input.Caller is WorksheetCaller.Adjustment or WorksheetCaller.AdjustmentAfterBordx && input.PayTo == "M");
        if (overrideApplies && Num(input.ApprovedOverride) is { } o && o > 0)
        {
            v[161] = o;
            v[162] = 0m;
        }
        else
        {
            v[161] = finalApproved;
            v[162] = totalFMem;
        }
        v[163] = finalBal - finalApproved;
        v[164] = finalBal;
        v[165] = 0m;
        v[166] = 0m;
        if (input.Caller != WorksheetCaller.MemberInvoice) v[468] = mcoFee;
        return v;
    }

    /// <summary>GetMin(X, Y) with a numeric Y: the smaller; 0 → "-".</summary>
    public static decimal? Min(decimal x, decimal y) => (x > y ? y : x) is var m && m == 0 ? null : m;

    /// <summary>GetMax(X, Y): the desktop returns Y in both numeric branches (reproduced); 0 → "-".</summary>
    public static decimal? Max(decimal x, decimal y) => y == 0 ? null : y;

    /// <summary>DueToFromMBM: cashless → approved − actual when the actual is higher; reimbursement → the approved amount; 0 → "-".</summary>
    public static decimal? DueToFromMember(string payTo, decimal actual, decimal approved) => payTo switch
    {
        "C" => actual > approved ? approved - actual : null,
        "M" => approved == 0 ? null : approved,
        _ => null,
    };

    /// <summary>DueToFromCLN: limit − amount when the amount is higher; 0 → "-".</summary>
    public static decimal? DueToFromClinic(decimal amount, decimal limit) => amount > limit ? limit - amount : null;

    private static decimal? Num(string? text) =>
        decimal.TryParse(text?.Replace(",", "").Trim(), System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : null;
}
