using System.Globalization;

namespace Cis.Application.Membership;

/// <summary>A claim row of CIS_FFS2018_test (cis.MbmEnq_Utilisation).</summary>
public sealed class EnqClaimRow
{
    /// <summary>OPTICAL / DENTAL / MATERNITY / DENTALEXTRACTION / GP / SP.</summary>
    public string? TreatmentType { get; init; }
    /// <summary>FFS (CLMMCOProcFee).</summary>
    public decimal? Ffs { get; init; }
    /// <summary>Clinic type G / S.</summary>
    public string? ClnType { get; init; }
    /// <summary>Maternity amount.</summary>
    public decimal? ClmMaternity { get; init; }
    /// <summary>Cover id.</summary>
    public string? CoverId { get; init; }
    /// <summary>Bordereaux amount.</summary>
    public decimal? ClmBordxAmt { get; init; }
    /// <summary>PLN.PLNGPSPShared.</summary>
    public string? PlnGpSpShared { get; init; }
    /// <summary>Others not deducted.</summary>
    public decimal? ClmOthersNoDed { get; init; }
    /// <summary>Typhoid jab not deducted (the desktop's "thyroid").</summary>
    public decimal? ClmTyphiadJabNoDed { get; init; }
    /// <summary>Dental amount.</summary>
    public decimal? ClmDental { get; init; }
    /// <summary>Dental / optical utilisation.</summary>
    public decimal? Utilization { get; init; }
}

/// <summary>A row of the CISDB dental procedures (tmpDlimit_local_test1 / tmpDentallimit_OneLimit).</summary>
public sealed class EnqDentalRow
{
    /// <summary>Cover id.</summary>
    public string? CoverId { get; init; }
    /// <summary>Treatment type.</summary>
    public string? TreatmentType { get; init; }
    /// <summary>Bordereaux amount (tmpDlimit_local_test1).</summary>
    public decimal? ClmBordxAmt { get; init; }
    /// <summary>Dental bordereaux amount (tmpDentallimit_OneLimit).</summary>
    public decimal? BordxDentalAmount { get; init; }
    /// <summary>tp1 (tmpDentallimit_OneLimit).</summary>
    public decimal? Tp1 { get; init; }
}

/// <summary>The plan's dental / optical / specialist flags that select the desktop's benefit scenario.</summary>
/// <param name="Pd">PDentalStatus.</param>
/// <param name="Sd">SDentalStatus.</param>
/// <param name="Gs">IsGPDentalShared.</param>
/// <param name="Os">IsDentalOpticalShared.</param>
/// <param name="Ps">PSpecStatus.</param>
/// <param name="Ss">SSpecStatus.</param>
/// <param name="Pdl">tmpPDentalAnnLimit as the desktop's string ("0.00" for NULL, "0" when no AnnualLimit row).</param>
public sealed record EnqFlags(string Pd, string Sd, string Gs, string Os, string Ps, string Ss, string Pdl)
{
    /// <summary>The six flags joined with "|".</summary>
    public string Key => $"{Pd}|{Sd}|{Gs}|{Os}|{Ps}|{Ss}";

    /// <summary>tmpPDentalAnnLimit = "" Or tmpPDentalAnnLimit = "0.00".</summary>
    public bool PdlBlank => Pdl is "" or "0.00";

    /// <summary>tmpPDentalAnnLimit > 0 (numeric comparison).</summary>
    public bool PdlPositive => MbmEnquiryRules.Num(Pdl) > 0;
}

/// <summary>Limit / utilised / balance texts of one benefit (the FPGridTopup columns).</summary>
/// <param name="Limit">Limit.</param>
/// <param name="Utilised">Utilised.</param>
/// <param name="Balance">Balance.</param>
public record EnqBenefit(string Limit, string Utilised, string Balance);

/// <summary>FrmMBMCLNEnq's calculations, rule for rule (scenario chains of showMBM, PrincipleDentallimits, tmppridentallimit,
/// Dentallimit_OneLimit and the vision labels).</summary>
public static class MbmEnquiryRules
{
    /// <summary>"N/A".</summary>
    public const string NA = "N/A";

    /// <summary>CStr of the AnnualLimit.PDentalAnnLimit read (IIf(Len(x), x, "0.00")); "0" when there is no row.</summary>
    public static string Pdl(bool rowFound, decimal? value) =>
        !rowFound ? "0" : value is null ? "0.00" : value.Value.ToString("0.############", CultureInfo.InvariantCulture);

    /// <summary>showMBM's principal utilisation (no specialist limit): the amount a claim row adds, or null when the row adds nothing.</summary>
    public static decimal? GpAdd(EnqFlags f, bool maternityPlan, EnqClaimRow r)
    {
        decimal m = r.ClmMaternity ?? 0, o = r.ClmOthersNoDed ?? 0, t = r.ClmTyphiadJabNoDed ?? 0, u = r.Utilization ?? 0, b = r.ClmBordxAmt ?? 0;
        var cov00 = (r.CoverId ?? "").Trim() == "00";
        decimal? Mi(decimal yes, decimal no) => b - (maternityPlan ? yes : no);
        return f.Key switch
        {
            "C||A|A|C|" => Mi(m + o, o),
            "C||A|B||" when f.PdlBlank => Mi(m + o + t, t + o),
            "C||A|C|Y|" => Mi(m + o + t, t + o),
            "C||B|A|C|" => Mi(o + m + t + u, o + t + u),
            "Y||B|A|Y|" => Mi(m + t + o + u, o + t + u),
            "C||B|B||" or "Y||B|B||" or "Y|Y|B|B||" => Mi(m + t + o + u, t + o + u),
            "Y||A|B||" when f.PdlBlank || f.Pdl == "" => Mi(m + o + t, t + o),
            "C||B|A||" => Mi(m + t + o, t + o),
            "C||D|B||" => Mi(m + t, m + t),
            "Y|C|A|A||C" => Mi(m + t + u, t + u),
            "||||||" => Mi(m + t + o, t + o),
            "C||A|B||" when f.PdlPositive => Mi(m + t + o, t),
            "C||A|A||C" => Mi(m + t + o, t),
            "Y|C|C|D|Y|" when cov00 => Mi(m + o + t + u, o + t + u),
            "||A|D|C|" => Mi(m + o + t, o + t),
            "||B|A|Y|" => (r.TreatmentType ?? "").Trim() == "OPTICAL" && cov00 ? b - (m + o + t) : null,
            "Y|Y|D|B|Y|" when f.PdlPositive => Mi(m + o + t + u, o + t + u),
            "Y|Y|D|B||" when f.PdlPositive => Mi(m + o + t + u, t + u),
            "Y|||E|Y|" when f.PdlPositive => Mi(m + o + t + u, t + u),
            "Y|C|E|A|Y|C" => Mi(m + o + t + u, t + u),
            "|||F|Y|" => Mi(m + t + u, t + u),
            _ => b - (m + t + o),
        };
    }

    /// <summary>showMBM with a specialist limit: GP utilisation of a row (null = not counted).</summary>
    public static decimal? GpAddWithSpec(bool maternityPlan, bool dentalPlan, EnqClaimRow r)
    {
        if (r.PlnGpSpShared?.Trim() != "Y" && (r.ClnType is null || r.ClnType.Trim() == "S")) return null;
        decimal m = r.ClmMaternity ?? 0, o = r.ClmOthersNoDed ?? 0, t = r.ClmTyphiadJabNoDed ?? 0, d = r.ClmDental ?? 0;
        return (r.ClmBordxAmt ?? 0) - (o + t + (maternityPlan ? m : 0) + (dentalPlan ? d : 0));
    }

    /// <summary>tmpPrioptiocalimitlimit: the principal's optical claims.</summary>
    public static decimal Optical(IEnumerable<EnqClaimRow> rows) =>
        rows.Where(r => r.TreatmentType?.Trim() == "OPTICAL" && r.CoverId?.Trim() == "00").Sum(r => r.ClmBordxAmt ?? 0);

    /// <summary>The vision (optical) labels of showMBM.</summary>
    public static EnqBenefit Vision(EnqFlags f, decimal optical, string opticalLimit)
    {
        var acc = Money(optical);
        var na = new EnqBenefit(NA, NA, NA);
        var accOnly = new EnqBenefit(NA, acc, NA);
        var limited = new EnqBenefit(opticalLimit, acc, Money(Num(opticalLimit) - R2(optical)));
        return f.Key switch
        {
            "C||A|A|C|" => accOnly,
            "C||A|B||" when f.PdlBlank => accOnly,
            "C||A|C|Y|" or "C||B|A|C|" or "Y||B|A|Y|" or "C||B|B||" => accOnly,
            "Y||B|B||" or "Y|Y|B|B||" => na,
            "Y||A|B||" => na,
            "C||B|A||" => accOnly,
            "C||D|B||" => na,
            "Y|C|A|A||C" => accOnly,
            "||||||" => na,
            "C||A|B||" when f.PdlPositive => na,
            "C||A|A||C" or "Y|C|C|D|Y|" or "||A|D|C|" or "||B|A|Y|" or "Y|Y|D|B|Y|" => accOnly,
            "Y|||E|Y|" when f.PdlPositive => limited,
            "|||F|Y|" or "Y|C|E|A|Y|C" => limited,
            _ => na,
        };
    }

    /// <summary>Which dental procedure PrincipleDentallimits uses: 0 none, 1 tmppridentallimit, 2 Dentallimit_OneLimit.</summary>
    public static int DentalSource(EnqFlags f) => f.Key switch
    {
        "C||A|A|C|" => 1,
        "C||A|B||" when f.PdlBlank => 1,
        "C||A|C|Y|" => 1,
        "C||B|A|C|" or "Y||B|A|Y|" or "C||B|B||" or "Y||B|B||" => 2,
        "Y|Y|B|B||" => 1,
        "Y||A|B||" when f.Pdl == "" => 2,
        "Y||A|B||" or "C||B|A||" or "C||D|B||" or "Y|C|A|A||C" => 1,
        "||||||" => 0,
        "C||A|B||" when f.PdlPositive => 1,
        "C||A|A||C" or "Y|C|C|D|Y|" => 1,
        "||A|D|C|" or "||B|A|Y|" => 0,
        "Y|Y|D|B|Y|" => 1,
        "Y|Y|D|B||" when f.PdlPositive => 0,
        "Y|C|E|A|Y|C" => 1,
        "Y|||E|Y|" when f.PdlPositive => 1,
        _ => 0,
    };

    /// <summary>tmppridentallimit: (last row's amount, total) of the rows its scenario counts.</summary>
    public static (decimal Last, decimal Total) PrincipalDental(EnqFlags f, IEnumerable<EnqDentalRow> rows)
    {
        decimal last = 0, total = 0;
        foreach (var r in rows)
        {
            if (!PrincipalDentalCounts(f, r)) continue;
            last = r.ClmBordxAmt ?? 0;
            total += last;
        }
        return (last, total);
    }

    private static bool PrincipalDentalCounts(EnqFlags f, EnqDentalRow r)
    {
        string tt = (r.TreatmentType ?? "").Trim(), spec = $"{f.Ps}|{f.Ss}";
        bool cov00 = (r.CoverId ?? "").Trim() == "00", dental = tt == "DENTAL", dentalExt = dental || tt == "DENTALEXTRACTION",
            all = dentalExt || tt == "DENTAL EXTRACTION (EXCLUDE IN DENTAL LIMIT)";
        var head = $"{f.Pd}|{f.Sd}|{f.Gs}|{f.Os}";
        if (head == "C||A|A" && cov00 && dental) return spec == "C|";
        if (head == "C||A|B" && cov00 && all) return spec == "|";
        if (head == "C||A|C" && cov00 && all) return spec == "Y|";
        if (head == "C||B|B" && cov00 && dentalExt) return spec == "|";
        if (head == "Y|Y|B|B" && cov00 && dentalExt) return spec == "|";
        if (head == "Y||A|B" && cov00 && all) return spec == "|";
        if (head == "Y||A|B" && all) return spec == "|";
        if (head == "C||B|A" && all) return spec == "|";
        if (head == "C||D|B" && all) return spec == "|";
        if (head == "Y|C|A|A" && dentalExt) return spec == "|C";
        if (head == "C||A|B" && all) return spec == "|";
        if (head == "C||A|A" && all) return spec == "|C";
        if (head == "Y|Y|D|B" && cov00 && dental && spec == "Y|") return true;
        if (f.Key == "Y|||E|Y|" && dental && f.PdlPositive) return true;
        if (f.Key == "Y|C|E|A|Y|C") return dental;
        if (head == "Y|C|C|D" && cov00 && dentalExt) return spec == "Y|";
        return false;
    }

    /// <summary>Dentallimit_OneLimit: (tmpaccudentallmt, tmpPrincipalDentalLimit, StrOneLimitUtli, strtp1).</summary>
    public static (decimal Acc, decimal Principal, decimal OneLimit, decimal Tp1) OneLimitDental(EnqFlags f, string asoType, IEnumerable<EnqDentalRow> rows)
    {
        decimal acc = 0, principal = 0, one = 0, tp1 = 0;
        var head = $"{f.Pd}|{f.Sd}|{f.Gs}|{f.Os}";
        var spec = $"{f.Ps}|{f.Ss}";
        foreach (var r in rows)
        {
            decimal b = r.BordxDentalAmount ?? 0;
            bool own = (r.CoverId ?? "").Trim() == "00" && (r.TreatmentType ?? "").Trim() == "DENTAL";
            if (head == "C||B|A")
            {
                if (spec != "C|") continue;
                if (own) acc += b; else one += b;
            }
            else if (head == "Y||B|A" && (r.TreatmentType ?? "").Trim() == "DENTAL")
            {
                if (spec != "Y|") continue;
                acc = b;
                principal += b;
            }
            else if (head == "C||B|B")
            {
                if (spec != "|") continue;
                if (own) acc += b; else one += b;
                if (asoType is "2" or "3") tp1 += r.Tp1 ?? 0;
            }
            else if (head is "Y||B|B" or "Y||A|B")
            {
                if (spec != "|") continue;
                acc = b;
                principal += b;
            }
        }
        return (acc, principal, one, tp1);
    }

    /// <summary>PrincipleDentallimits' labels from the dental procedure totals.</summary>
    public static EnqBenefit DentalLabels(EnqFlags f, string asoType, (decimal Last, decimal Total) pri, (decimal Acc, decimal Principal, decimal OneLimit, decimal Tp1) one,
        decimal optical)
    {
        var na = new EnqBenefit(NA, NA, NA);
        decimal lim = Num(f.Pdl);
        EnqBenefit AccOnly() => new(NA, Money(pri.Total), NA);
        return f.Key switch
        {
            "C||A|A|C|" => AccOnly(),
            "C||A|B||" when f.PdlBlank => AccOnly(),
            "C||A|C|Y|" => AccOnly(),
            "C||B|A|C|" => new(Money(lim), Money(one.Acc), Money(lim - one.OneLimit - one.Acc)),
            "Y||B|A|Y|" => new(Money(lim), Money(one.Principal), Money(R2(lim) - R2(one.Principal) - R2(optical))),
            "C||B|B||" when asoType is "2" or "3" => new(Money(lim), Money(one.Acc), Money(R2(lim) - R2(one.Tp1))),
            "C||B|B||" or "Y||B|B||" => new(Money(lim), Money(one.Acc), Money(R2(lim) - R2(one.Acc))),
            "Y|Y|B|B||" => new(Money(lim), Money(pri.Last), Money(R2(lim) - R2(pri.Last))),
            "Y||A|B||" when f.Pdl == "" => new(Money(lim), Money(one.Principal), Money(R2(lim) - R2(one.Acc))),
            "Y||A|B||" or "C||B|A||" or "C||D|B||" or "Y|C|A|A||C" => AccOnly(),
            "||||||" => na,
            "C||A|B||" when f.PdlPositive => AccOnly(),
            "C||A|A||C" => AccOnly(),
            "Y|C|C|D|Y|" or "Y|Y|D|B|Y|" => new(Money(lim), Money(pri.Total), Money(R2(lim) - R2(pri.Total) - R2(optical))),
            "||A|D|C|" or "||B|A|Y|" => na,
            "Y|Y|D|B||" when f.PdlPositive => na,
            "Y|C|E|A|Y|C" => AccOnly(),
            "Y|||E|Y|" when f.PdlPositive => new(Money(lim), Money(pri.Last), Money(R2(lim) - R2(pri.Last))),
            _ => na,
        };
    }

    /// <summary>showMBMCov's claim filter of a cover: 2 = covers other than 00, 3 = this cover, null = none. <paramref name="specBranch"/>:
    /// with a specialist limit the desktop needs cover 01 for B.</summary>
    public static int? CoverMode(string status, string coverId, bool specBranch) => status switch
    {
        "B" when !specBranch || coverId == "01" => 2,
        "C" or "F" or "" => 3,
        "D" when coverId == "01" => 2,
        "E" => 3,
        _ => null,
    };

    /// <summary>showMBMCov (no specialist limit): a row's deduction of Utilization in the cover's accumulated amount.</summary>
    public static bool CoverSubtractsUtilisation(EnqFlags f) => f.Key switch
    {
        "C||A|C|Y|" or "C||B|A|C|" or "C||B|B||" or "Y|Y|B|B||" or "Y|C|E|A|Y|C" => true,
        "Y|Y|D|B||" => f.PdlPositive,
        _ => false,
    };

    /// <summary>showMBMCov (no specialist limit): (FFS, accumulated) of a cover; limit status A gives zero, F rows add nothing.</summary>
    public static (decimal Ffs, decimal Acc) CoverGp(EnqFlags f, string status, IEnumerable<EnqClaimRow> rows)
    {
        if (status == "A") return (0, 0);
        decimal ffs = 0, acc = 0;
        var sub = CoverSubtractsUtilisation(f);
        foreach (var r in rows)
        {
            if (status != "F")
            {
                ffs += r.Ffs ?? 0;
                acc += (r.ClmBordxAmt ?? 0) - (r.ClmOthersNoDed ?? 0) - (r.ClmMaternity ?? 0);
            }
            if (sub) acc -= r.Utilization ?? 0;
        }
        return (ffs, acc);
    }

    /// <summary>showMBMCov with a specialist limit: GP (FFS, accumulated) of a cover from its non-specialist rows.</summary>
    public static (decimal Ffs, decimal Acc) CoverGpWithSpec(string status, IEnumerable<EnqClaimRow> rows)
    {
        if (status == "A") return (0, 0);
        var gp = rows.Where(r => r.ClnType is not null && r.ClnType.Trim() != "S").ToList();
        return (gp.Sum(r => r.Ffs ?? 0), gp.Sum(r => (r.ClmBordxAmt ?? 0) - (r.ClmOthersNoDed ?? 0)));
    }

    /// <summary>showMBMCov: specialist (FFS, accumulated) of a cover; the desktop tests the GP status for blank in its second branch.</summary>
    public static (decimal Ffs, decimal Acc) CoverSpec(string specStatus, string gpStatus, string coverId, IEnumerable<EnqClaimRow> rows)
    {
        if (specStatus == "A") return (0, 0);
        bool counts = (specStatus == "B" && coverId == "01") || specStatus == "C" || gpStatus == "" || (specStatus == "D" && coverId == "01");
        if (!counts) return (0, 0);
        var sp = rows.Where(r => r.ClnType?.Trim() == "S").ToList();
        return (sp.Sum(r => r.Ffs ?? 0), sp.Sum(r => (r.ClmBordxAmt ?? 0) - (r.ClmOthersNoDed ?? 0) - (r.ClmMaternity ?? 0)));
    }

    private static bool SuppDentalKey(EnqFlags f) => f.Key switch
    {
        "C||A|A|C|" or "C||A|C|Y|" or "C||B|A|C|" or "Y||B|A|Y|" or "C||B|B||" or "Y||B|B||" or "Y|Y|B|B||" or "Y||A|B||" or "C||B|A||" or "C||D|B||"
            or "Y|C|A|A||C" or "C||A|A||C" or "Y|C|C|D|Y|" => true,
        "C||A|B||" => f.PdlBlank || f.PdlPositive,
        "Y|Y|D|B||" => f.PdlPositive,
        _ => false,
    };

    /// <summary>tmpsuppdentallimit: (last row's amount, total) of a cover's dental rows.</summary>
    public static (decimal Last, decimal Total) SuppDental(EnqFlags f, IEnumerable<EnqDentalRow> rows)
    {
        decimal last = 0, total = 0;
        foreach (var r in rows)
        {
            string cov = (r.CoverId ?? "").Trim(), tt = (r.TreatmentType ?? "").Trim();
            if (tt != "DENTAL") continue;
            if (!((SuppDentalKey(f) && cov != "00") || (f.Key == "Y|Y|D|B|Y|" && cov == "01") || f.Key == "Y|C|E|A|Y|C")) continue;
            last = r.ClmBordxAmt ?? 0;
            total += last;
        }
        return (last, total);
    }

    /// <summary>Suppoptiocalimitlimit: a cover's optical claims.</summary>
    public static decimal SuppOptical(EnqFlags f, IEnumerable<EnqDentalRow> rows) => rows.Where(r =>
    {
        string cov = (r.CoverId ?? "").Trim(), tt = (r.TreatmentType ?? "").Trim();
        if (f.Key == "Y|C|E|A|Y|C") return true;
        if (f.Key == "Y|Y|D|B|Y|") return cov == "01" && tt == "OPTICAL";
        var key = f.Key is "C||A|A|C|" or "C||A|C|Y|" or "C||B|A|C|" or "C||B|A||" or "Y|C|A|A||C" or "C||A|A||C" or "Y|C|C|D|Y|" or "||A|D|C|"
                  || (f.Key == "C||A|B||" && f.PdlBlank);
        return key && cov != "00" && tt == "OPTICAL";
    }).Sum(r => r.ClmBordxAmt ?? 0);

    /// <summary>Which cover the optical figure of Supplyoptiocallimit is computed for (null = N/A).</summary>
    public static string? SuppOpticalCover(EnqFlags f, string coverId) => f.Key switch
    {
        "C||A|A|C|" or "C||A|C|Y|" or "C||B|A|C|" or "C||B|A||" or "Y|C|A|A||C" or "C||A|A||C" or "Y|C|C|D|Y|" or "Y|C|E|A|Y|C" => coverId,
        "C||A|B||" when f.PdlBlank => coverId,
        "Y|Y|D|B|Y|" => coverId == "01" ? "01" : null,
        _ => null,
    };

    /// <summary>Whether SupplyDentallimits reads the cover's dental rows.</summary>
    public static bool SuppDentalRead(EnqFlags f) => f.Key switch
    {
        "C||A|A|C|" or "C||A|C|Y|" or "C||B|A|C|" or "C||B|B||" or "Y|Y|B|B||" or "C||B|A||" or "C||D|B||" or "Y|C|A|A||C" or "Y|C|C|D|Y|" or "||A|D|C|"
            or "Y|C|E|A|Y|C" or "Y|Y|D|B|Y|" => true,
        "C||A|B||" => f.PdlBlank || f.PdlPositive,
        _ => false,
    };

    /// <summary>SupplyDentallimits' labels of a cover.</summary>
    public static EnqBenefit SuppDentalLabels(EnqFlags f, string coverId, decimal sDentalLimit, (decimal Last, decimal Total) d, decimal suppOptical)
    {
        var na = new EnqBenefit(NA, NA, NA);
        var total = new EnqBenefit(NA, Money(d.Total), NA);
        return f.Key switch
        {
            "C||A|A|C|" or "C||A|C|Y|" or "C||B|A|C|" or "C||B|B||" or "C||B|A||" or "C||D|B||" or "Y|C|A|A||C" or "Y|C|C|D|Y|" or "||A|D|C|" => total,
            "C||A|B||" when f.PdlBlank || f.PdlPositive => total,
            "Y|Y|B|B||" => new(Money(sDentalLimit), Money(d.Last), Money(sDentalLimit - d.Last)),
            "Y|C|E|A|Y|C" => new(NA, Money(d.Last), NA),
            "Y|Y|D|B|Y|" when coverId == "01" => new(Money(sDentalLimit), Money(d.Last), Money(sDentalLimit - (d.Last + R2(suppOptical)))),
            _ => na,
        };
    }

    /// <summary>VB CDbl of a label ("1,234.00" → 1234; non-numeric → 0).</summary>
    public static decimal Num(string? text) =>
        decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;

    /// <summary>True when VB's IsNumeric accepts the text.</summary>
    public static bool IsNum(string? text) => decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out _);

    /// <summary>Format(x, "#,##0.00").</summary>
    public static string Money(decimal v) => v.ToString("#,##0.00", CultureInfo.InvariantCulture);

    /// <summary>Format(label, "#,##0.00"): numbers are formatted, texts (FULL, N/A) kept.</summary>
    public static string Money(string label) => IsNum(label) ? Money(Num(label)) : label;

    /// <summary>Round half away from zero to 2 decimals (VB Format of an amount).</summary>
    public static decimal R2(decimal v) => Math.Round(v, 2, MidpointRounding.AwayFromZero);
}
