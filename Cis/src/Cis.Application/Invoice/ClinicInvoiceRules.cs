using System.Globalization;

namespace Cis.Application.Invoice;

/// <summary>A claim of CIS_FFS2013 (cis.InvClinic_Claims).</summary>
public sealed record InvClaim
{
    /// <summary>Claim.CLMMCOProcFee.</summary>
    public decimal? Ffs { get; init; }
    /// <summary>CLN.CLNType.</summary>
    public string? ClnType { get; init; }
    /// <summary>Maternity.</summary>
    public decimal ClmMaternity { get; init; }
    /// <summary>Cover.</summary>
    public string? CoverId { get; init; }
    /// <summary>Bordereaux amount (total actual less co-pay).</summary>
    public decimal? ClmBordxAmt { get; init; }
    /// <summary>PLN.PLNGPSPShared.</summary>
    public string? PlnGpSpShared { get; init; }
    /// <summary>Others without deduction (or the BTBG amount).</summary>
    public decimal ClmOthersNoDed { get; init; }
    /// <summary>Typhoid jab.</summary>
    public decimal? ClmTyphiadJabNoDed { get; init; }
    /// <summary>Dental.</summary>
    public decimal ClmDental { get; init; }
}

/// <summary>A medication line of the 7-line grid.</summary>
/// <param name="Type">Medication / injection.</param>
/// <param name="Days">Days.</param>
/// <param name="Qty">Quantity.</param>
/// <param name="Rate">Price.</param>
public record InvoiceMedication(string? Type, string? Days, string? Qty, string? Rate);

/// <summary>The case typed on frmInvoice (header, member, visit, items). A holds A1..A33 by index (A1–A6 are computed from B), B holds B1..B8.</summary>
public sealed record ClinicInvoiceEntry
{
    /// <summary>Invoice header.</summary>
    public string? InvoiceNo { get; init; }
    /// <summary>Invoice date dd/mm/yyyy.</summary>
    public string? InvoiceDate { get; init; }
    /// <summary>Invoice receipt date.</summary>
    public string? ReceiptDate { get; init; }
    /// <summary>Receipt mode (code or "code - name").</summary>
    public string? ReceiptMode { get; init; }
    /// <summary>Clinic code.</summary>
    public string? ClinicCode { get; init; }
    /// <summary>Total invoice amount (TxtTotInvoiceAmt).</summary>
    public string? TotalInvoiceAmount { get; init; }
    /// <summary>Member and cover.</summary>
    public string? MbmNumber { get; init; }
    /// <summary>Cover id.</summary>
    public string? CoverId { get; init; }
    /// <summary>Visit date / time, bill no., MC days.</summary>
    public string? VisitDate { get; init; }
    /// <summary>hh:mm.</summary>
    public string? VisitTime { get; init; }
    /// <summary>Bill no.</summary>
    public string? BillNo { get; init; }
    /// <summary>MC days.</summary>
    public decimal? McDays { get; init; }
    /// <summary>Diagnosis 1–5 codes.</summary>
    public IReadOnlyList<string?> Diagnoses { get; init; } = [];
    /// <summary>Diagnosis 1–5 descriptions.</summary>
    public IReadOnlyList<string?> DiagnosisDescriptions { get; init; } = [];
    /// <summary>Others diagnosis (diagnosis 5 with the "others" indicator).</summary>
    public string? DiagnosisOther { get; init; }
    /// <summary>A1..A33 (index 1..33; index 0 unused).</summary>
    public IReadOnlyList<decimal?> A { get; init; } = [];
    /// <summary>B1..B8 (index 1..8).</summary>
    public IReadOnlyList<decimal?> B { get; init; } = [];
    /// <summary>Clinic bill amount (txtCLNBillAmt).</summary>
    public string? ClinicBillAmount { get; init; }
    /// <summary>FFS Y / N.</summary>
    public string? Ffs { get; init; }
    /// <summary>Process month MM/YYYY.</summary>
    public string? ProcessMonth { get; init; }
    /// <summary>Claimability A / M / R.</summary>
    public string? Claimability { get; init; }
    /// <summary>Remark.</summary>
    public string? Remark { get; init; }
    /// <summary>Adjusted annual limit (txtMBMAdjAnnLmt).</summary>
    public decimal? AdjustedAnnualLimit { get; init; }
    /// <summary>Total excess (txtExcess).</summary>
    public decimal? Excess { get; init; }
    /// <summary>Limit exceeded.</summary>
    public decimal? LimitExceeded { get; init; }
    /// <summary>Procedure / laboratory / immunisation "others" descriptions.</summary>
    public string? POthDesc { get; init; }
    /// <summary>Laboratory others.</summary>
    public string? LOthDesc { get; init; }
    /// <summary>Immunisation others.</summary>
    public string? IOthDesc { get; init; }
    /// <summary>Medication lines (up to 7).</summary>
    public IReadOnlyList<InvoiceMedication> Medications { get; init; } = [];
}

/// <summary>frmInvoice's calculations that need no database.</summary>
public static class ClinicInvoiceRules
{
    /// <summary>getTotal: A1..A6 from the GP (B1–B4) and SP (B5–B8) fees.</summary>
    public static decimal[] Amounts(ClinicInvoiceEntry e)
    {
        var a = new decimal[34];
        for (var i = 1; i <= 33; i++) a[i] = i < e.A.Count ? e.A[i] ?? 0 : 0;
        decimal B(int i) => i < e.B.Count ? e.B[i] ?? 0 : 0;
        var c1 = B(1) + B(2) + B(3) + B(4);
        var c2 = B(5) + B(6) + B(7) + B(8);
        a[1] = a[2] = a[3] = a[4] = a[5] = a[6] = 0;
        if (B(2) > 0) a[3] = c1; else if (B(3) > 0 && B(4) > 0) a[2] = c1; else a[1] = c1;
        if (B(6) > 0) a[6] = c2; else if (B(7) > 0 && B(8) > 0) a[5] = c2; else a[4] = c2;
        return a;
    }

    /// <summary>getTotal's IntInvAmt: every A amount except the co-pay (A29).</summary>
    public static decimal Total(decimal[] a) => Math.Round(Enumerable.Range(1, 33).Where(i => i != 29).Sum(i => a[i]), 2, MidpointRounding.AwayFromZero);

    /// <summary>The worksheet's item amounts in grid order (BNF 101–106, 301–309, 401–405, 501–505, 601–604, 701–702).</summary>
    public static decimal?[] GridItems(decimal[] a) =>
        [a[1], a[2], a[3], a[4], a[5], a[6], .. Enumerable.Range(10, 19).Select(i => (decimal?)a[i]), a[7], a[8], a[9], a[30], a[32], a[33]];

    /// <summary>CmdGo's utilisation: the claims' bordereaux amounts less typhoid, others without deduction and, by the plan's flags,
    /// maternity / dental. With a specialist limit, specialist-clinic claims go to the specialist total unless GP / SP is shared.</summary>
    public static (decimal Accum, decimal Specialist) Utilisation(IEnumerable<InvClaim> claims, bool maternityPlan, bool dentalPlan, bool specBranch)
    {
        decimal accum = 0, spec = 0;
        foreach (var c in claims)
        {
            var amount = c.ClmBordxAmt is { } b
                ? b - (c.ClmOthersNoDed + (c.ClmTyphiadJabNoDed ?? 0) + (maternityPlan ? c.ClmMaternity : 0) + (dentalPlan ? c.ClmDental : 0))
                : 0;
            if (specBranch && c.PlnGpSpShared?.Trim() != "Y" && c.ClnType?.Trim() == "S") spec += amount;
            else accum += amount;
        }
        return (accum, spec);
    }

    /// <summary>The CIS_FFS2013 filter by limit status: A the member, B / D the principal or the covers, C / E / F / blank the cover, else none.</summary>
    public static int? ClaimMode(string limitStatus, string covId) => limitStatus switch
    {
        "A" => 0,
        "B" or "D" => covId == "00" ? 1 : 2,
        "C" or "" or "E" or "F" => 3,
        _ => null,
    };

    /// <summary>ChkData (after the visit-date checks): the first desktop message, or null.</summary>
    public static string? Check(ClinicInvoiceEntry e, decimal[] a, decimal total, DateTime? eff, DateTime? exp, bool otherDiagnosis, DateTime today)
    {
        var visit = Date(e.VisitDate);
        if (string.IsNullOrWhiteSpace(e.MbmNumber)) return "Please keyin Member No ! ";
        if (string.IsNullOrWhiteSpace(e.VisitDate)) return "Please Enter Visit Date! ";
        if (visit is null) return "Invalid Visit date format!";
        if (visit > today) return "Visit Date cannot greater than system date ! ";
        if (eff is null) return "Invalid Effective date format!";
        if (exp is null) return "Invalid Expiry date format!";
        if (visit < eff) return "Visit Date cannot less than Membership Effective date ! ";
        if (visit > exp) return "Visit Date cannot greater than Membership Expiry date ! ";
        if (string.IsNullOrWhiteSpace(e.ClinicCode)) return "Please keyin the Clinic code ! ";
        if (string.IsNullOrWhiteSpace(e.InvoiceNo)) return "Please Enter Invoice No ! ";
        if (string.IsNullOrWhiteSpace(e.InvoiceDate)) return "Please Enter Invoice Date! ";
        if (string.IsNullOrWhiteSpace(e.ReceiptDate)) return "Please Enter Invoice Receipt Date ! ";
        if (string.IsNullOrWhiteSpace(e.TotalInvoiceAmount)) return "Please Keyin Invoice Amount ! ";
        if (total == 0) return "Total Invoice Amount cannot be Zero ! ";
        if (otherDiagnosis && string.IsNullOrWhiteSpace(e.DiagnosisOther)) return "Please Keyin Others Diagnosis ! ";
        if (string.IsNullOrWhiteSpace(e.ClinicBillAmount)) return "Please enter Bill Amount! ";
        // txtCLNBillAmt <> lblCLNBillAmt + A29, where lblCLNBillAmt = the total less A29.
        if (Num(e.ClinicBillAmount) != total) return "Please check Bill Amount! ";
        if (string.IsNullOrWhiteSpace(e.Claimability)) return "Please Enter Claimability ";
        return null;
    }

    /// <summary>txtVisitDate_LostFocus: the policy / cancellation checks that disable Add.</summary>
    public static string? VisitCheck(DateTime visit, DateTime today, bool principal, string policyStatus, DateTime? exp, string cancelDate, string coverCancelDate,
        string coverClinicExpDate)
    {
        const string Terminated = "Membership has been terminated before the entered visit date.",
            Range = "Visit date entered is not in range of Policy Effective Date and Expiry Date";
        if (visit > today) return "Visit Date cannot greater than system date ! ";
        var active = policyStatus == "A";
        if (principal && !active && Date(cancelDate) is { } c) return visit > c ? Terminated : null;
        if (principal && !active && visit > exp) return Range;
        if (!principal && !active && Date(coverCancelDate) is { } cc) return visit > cc ? Terminated : null;
        if (!principal && !active && Date(coverClinicExpDate) is { } ce) return visit > ce ? Range : null;
        if (active && principal && visit > exp) return Range;
        if (active && !principal && Date(coverClinicExpDate) is { } ce2 && visit > ce2) return Range;
        return null;
    }

    /// <summary>cmdVerify: the error list of the line (message only; the desktop also lists the member columns).</summary>
    public static List<string> Verify(ClinicInvoiceEntry e, decimal total, DateTime? eff, DateTime? exp, string balance, string visitLimit,
        IEnumerable<string> notPayableDiagnoses, DateTime today)
    {
        var errors = notPayableDiagnoses.Select(d => $"Diagnosis '{d}' is not payable ! ").ToList();
        if (string.IsNullOrWhiteSpace(e.InvoiceNo)) errors.Add("Invoice No Can't be null ! ");
        if (Date(e.InvoiceDate) is { } inv && eff is not null && exp is not null)
        {
            if (inv < eff) errors.Add("Invoice Date MUST BE GREATER than Clinical Effective date ! ");
            if (inv > exp.Value.AddDays(14)) errors.Add("Invoice date IS GREATER THAN Clinical Expiry date + 14 days ! ");
            if ((today - inv).TotalDays > 45) errors.Add("45 Days LATE ! ");
        }
        if (Date(e.VisitDate) is { } v && eff is not null && exp is not null)
        {
            if (v < eff) errors.Add("Visit date MUST BE AFTER the Effective date ! ");
            if (v > exp.Value.AddDays(14)) errors.Add("Visit date IS GREATER THAN Expiry date + 14 days ! ");
        }
        if (NumOrNull(balance) is { } bal && total > bal) errors.Add("Bill Amt > Balance Limit ! ");
        if (NumOrNull(visitLimit) is { } vl && total > vl) errors.Add("Bill Amt > Visit Limit ! ");
        // CInt rounds half to even.
        if (Math.Round(total, MidpointRounding.ToEven) != Math.Round(Num(e.TotalInvoiceAmount), MidpointRounding.ToEven))
            errors.Add("Invoice Amt is different from Total Bill Amt");
        return errors;
    }

    /// <summary>dd/mm/yyyy (also dd/MM/yyyy HH:mm stored dates), or null.</summary>
    public static DateTime? Date(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    /// <summary>A number of the form (commas allowed), else 0.</summary>
    public static decimal Num(string? text) => NumOrNull(text) ?? 0;

    /// <summary>A number of the form, or null.</summary>
    public static decimal? NumOrNull(string? text) =>
        decimal.TryParse(text?.Replace(",", "").Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;
}
