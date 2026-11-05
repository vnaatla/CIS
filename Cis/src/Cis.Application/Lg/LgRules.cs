using System.Globalization;
using Cis.Application.Invoice;

namespace Cis.Application.Lg;

/// <summary>GetLGAnnLimitSpecLimit + the plan inputs CheckMBMLimit reads.</summary>
public sealed record LgLimitInput
{
    /// <summary>The cover clicked (00 = principal).</summary>
    public string CoverId { get; init; } = "00";
    /// <summary>GetLGAnnLimitSpecLimit row (texts as the procedure returns them).</summary>
    public IReadOnlyDictionary<string, object?> Row { get; init; } = new Dictionary<string, object?>();
    /// <summary>CLNLG1 ASOInd.</summary>
    public bool Aso { get; init; }
    /// <summary>ASO annual limit (combined ? ASO : OP limit) and utilisation (combined ? IP + OP : OP).</summary>
    public (decimal Limit, decimal Utilised)? AsoValues { get; init; }
    /// <summary>MBMCISTopupHistory SP_TopupAmt total.</summary>
    public decimal SpTopup { get; init; }
    /// <summary>AnnualLimit.SuppLimitStatus of the plan (null = no row).</summary>
    public string? LimitStatus { get; init; }
    /// <summary>GetSpecLimit ("FULL", "N/A", an amount or "").</summary>
    public string SpecLimit { get; init; } = "";
    /// <summary>PLN.PLNMaternityInd = Y.</summary>
    public bool Maternity { get; init; }
    /// <summary>PLN.PDentalStatus = Y.</summary>
    public bool Dental { get; init; }
    /// <summary>PLN.IsPrinDepGPShared.</summary>
    public bool PrinDepGpShared { get; init; }
}

/// <summary>The limit labels of frmLGGeneration (texts as the desktop shows them).</summary>
public sealed record LgLimits
{
#pragma warning disable CS1591
    public string AnnualLimit { get; init; } = "";
    public string Accumulated { get; init; } = "";
    public string Balance { get; init; } = "";
    public string SpecAnnual { get; init; } = "";
    public string SpecAccumulated { get; init; } = "";
    public string SpecBalance { get; init; } = "";
    public string TopUp { get; init; } = "";
    public string SpTopUp { get; init; } = "";
    public string NoAnnLimitDed { get; init; } = "";
    public string NoSpecLimitDed { get; init; } = "";
    public string PlanCode { get; init; } = "";
    public string PlnGlNo { get; init; } = "";
    public string LgIssuedInd { get; init; } = "";
    public string LgIssuedAfterLimitExInd { get; init; } = "";
    public string UsingLimit { get; init; } = "";
    public string UsingLimitDesc { get; init; } = "";
    public string PolicyEffDate { get; init; } = "";
    /// <summary>GetLGAnnLimitSpecLimit HRApproval = Y (the desktop asked " All MO/SP GL Request Must be Come From HR." OK / Cancel).</summary>
    public bool HrApproval { get; init; }
#pragma warning restore CS1591
}

/// <summary>frmLGGeneration's calculations that need no database.</summary>
public static class LgRules
{
    /// <summary>The HR approval question.</summary>
    public const string HrQuestion = " All MO/SP GL Request Must be Come From HR.";

    /// <summary>VB IsNumeric on a trimmed text.</summary>
    public static bool IsNum(string? v) => decimal.TryParse(v?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out _);

    /// <summary>CDbl of a numeric text.</summary>
    public static decimal Num(string? v) => decimal.TryParse(v?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;

    /// <summary>Format(x, "#,##0.00") (texts that are not numbers stay as they are).</summary>
    public static string Fmt(string? v) => IsNum(v) ? Num(v).ToString("#,##0.00", CultureInfo.InvariantCulture) : (v ?? "").Trim();

    private static string Fmt(decimal v) => v.ToString("#,##0.00", CultureInfo.InvariantCulture);

    private static string T(IReadOnlyDictionary<string, object?> r, string k) => r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    /// <summary>CheckMBMLimit's claim filter: null = no claims read, else the cis.InvClinic_Claims mode (0 member, 1 cover 00, 2 other covers, 3 one cover).</summary>
    public static int? ClaimMode(string limitStatus, string covId, bool prinDepShared, bool specBranch, string specSuppStatus)
    {
        if (prinDepShared) return 0;
        return limitStatus switch
        {
            "A" => specBranch && specSuppStatus == "C" ? 3 : 0,
            "B" or "D" => covId == "00" ? 1 : 2,
            "C" or "" or "E" => 3,
            _ => null,
        };
    }

    /// <summary>CheckMBMLimit: the labels from GetLGAnnLimitSpecLimit, then the utilisation recomputed from the claims (CIS_FFS2013).</summary>
    /// <param name="i">Inputs.</param>
    /// <param name="claims">Claims of <see cref="ClaimMode"/> (empty when it is null).</param>
    public static LgLimits Limits(LgLimitInput i, IReadOnlyList<InvClaim> claims)
    {
        var r = i.Row;
        var cov = i.CoverId;
        var status = i.LimitStatus is null ? "" : i.LimitStatus.Trim() is { Length: > 0 } s ? s : "A";
        var specSupp = T(r, "SpecSuppLimitStatus");
        var suppAnn = Fmt(T(r, "SuppLimit"));
        var ann = Fmt(T(r, "OriAnnLimit"));
        var specAnn = Fmt(T(r, "OriSpecAnnLimit"));
        var topUp = Fmt(T(r, "TopupAmt"));
        var lgIssued = cov != "00" ? (T(r, "LGIssuedSuppInd") == "N" ? "N" : "Y") : (T(r, "LGIssuedInd") is { Length: > 0 } li ? li : "N");
        string tmpAnn = i.Aso ? Fmt((i.AsoValues?.Limit ?? 0).ToString(CultureInfo.InvariantCulture)) : "";
        var utilised = i.AsoValues?.Utilised ?? 0;
        var tmpSpec = i.Aso ? "" : (IsNum(i.SpecLimit) ? Fmt(i.SpecLimit) : i.SpecLimit);
        if (status == "E" && cov == "00") tmpAnn = "FULL";
        if (status == "E" && cov != "00") ann = suppAnn;

        decimal accum = 0, specAcc = 0, noAnnDed = 0;
        var specBranch = specAnn.Trim() != "N/A";
        foreach (var c in claims)
        {
            var ded = c.ClmOthersNoDed + (c.ClmTyphiadJabNoDed ?? 0) + (i.Maternity ? c.ClmMaternity : 0) + (i.Dental ? c.ClmDental : 0);
            var amount = (c.ClmBordxAmt ?? 0) - ded;
            var toSpec = false;
            if (specBranch && c.PlnGpSpShared?.Trim() != "Y" && c.ClnType?.Trim() == "S")
            {
                var cc = c.CoverId?.Trim() ?? "";
                var counts = status switch
                {
                    "A" => specSupp != "C" || cc == cov,
                    "B" or "D" => cov == "00" ? cc == "00" : cc != "00",
                    "C" or "" or "E" => cc == cov,
                    _ => false,
                };
                if (!counts) continue;
                toSpec = true;
            }
            if (toSpec) specAcc += amount; else accum += amount;
            noAnnDed += ded;
        }
        if (i.Aso) { accum = utilised; ann = tmpAnn; }

        var bal = IsNum(ann) ? Fmt(Num(ann) + (IsNum(topUp) ? Num(topUp) : 0) - accum) : ann;
        string sAnn, sAcc, sBal;
        if (!specBranch) { sAnn = tmpSpec; sAcc = "N/A"; sBal = "N/A"; }
        else sAnn = tmpSpec;
        if (IsNum(sAnn))
        {
            sAcc = Fmt(specAcc);
            sBal = IsNum(tmpSpec) ? Fmt(Num(tmpSpec) + i.SpTopup - specAcc) : "FULL";
        }
        else if (sAnn.Trim() == "FULL") { sAcc = Fmt(specAcc); sBal = "FULL"; }
        else { sAnn = sAnn.Trim() == "N/A" ? sAnn : "N/A"; sAcc = "N/A"; sBal = "N/A"; }
        if ((status == "C" || status == "E") && cov == "00" && ((IsNum(ann) && Num(ann) == 0) || ann == "FULL")) bal = "FULL";
        var accText = Fmt(accum);
        if (i.PrinDepGpShared) { ann = "N/A"; bal = "N/A"; accText = "N/A"; }
        return new LgLimits
        {
            AnnualLimit = Fmt(ann), Accumulated = accText, Balance = Fmt(bal), SpecAnnual = sAnn, SpecAccumulated = sAcc, SpecBalance = sBal,
            TopUp = topUp, SpTopUp = Fmt(i.SpTopup), NoAnnLimitDed = Fmt(noAnnDed), NoSpecLimitDed = Fmt(T(r, "SNoLimitDedTotal")), PlanCode = T(r, "PLNCode"),
            PlnGlNo = T(r, "PLNGLNo"), LgIssuedInd = lgIssued, LgIssuedAfterLimitExInd = T(r, "LGIssuedAfterLimitExInd") is { Length: > 0 } ae ? ae : "N",
            UsingLimit = T(r, "LGUsingLimit") is { Length: > 0 } ul ? ul : "N/A", UsingLimitDesc = T(r, "LGUsingLimitDesc") is { Length: > 0 } ud ? ud : "N/A",
            PolicyEffDate = T(r, "PolEffDate"), HrApproval = T(r, "HRApproval") == "Y",
        };
    }

    /// <summary>lstLG_Click: the labels straight from GetLGAnnLimitSpecLimit (no recalculation; LGIssuedInd always the principal's, as the second
    /// assignment overwrites the supplementary one). ASO members keep the utilisation of the last member click.</summary>
    public static LgLimits RawLimits(IReadOnlyDictionary<string, object?> r, LgLimits? previous, bool aso) => new()
    {
        AnnualLimit = Fmt(T(r, "OriAnnLimit")), Accumulated = aso ? previous?.Accumulated ?? "" : Fmt(T(r, "GUtilized")), Balance = Fmt(T(r, "GBalance")),
        SpecAnnual = Fmt(T(r, "OriSpecAnnLimit")), SpecAccumulated = Fmt(T(r, "SUtilized")), SpecBalance = Fmt(T(r, "SBalance")), TopUp = Fmt(T(r, "TopupAmt")),
        SpTopUp = previous?.SpTopUp ?? "", NoAnnLimitDed = Fmt(T(r, "GNoLimitDedTotal")), NoSpecLimitDed = Fmt(T(r, "SNoLimitDedTotal")), PlanCode = T(r, "PLNCode"),
        PlnGlNo = T(r, "PLNGLNo"), LgIssuedInd = T(r, "LGIssuedInd") is { Length: > 0 } li ? li : "N",
        LgIssuedAfterLimitExInd = T(r, "LGIssuedAfterLimitExInd") is { Length: > 0 } ae ? ae : "N",
        UsingLimit = T(r, "LGUsingLimit") is { Length: > 0 } ul ? ul : "N/A", UsingLimitDesc = T(r, "LGUsingLimitDesc") is { Length: > 0 } ud ? ud : "N/A",
        PolicyEffDate = T(r, "PolEffDate"), HrApproval = false,
    };

    /// <summary>ValidateLG: whether an LG may be issued (IsValidLg) from SP_ValidateLG1 and the balances.</summary>
    public static bool IsValidLg(string lgIssuedInd, string gpSpShared, string afterLimitExInd, string balance, string specBalance)
    {
        bool? valid = lgIssuedInd == "Y" ? true : lgIssuedInd == "N" ? false : null;
        if (gpSpShared == "Y")
            valid = balance.Trim() == "FULL" || (IsNum(balance) && Num(balance) > 0);
        else if (specBalance.Length > 0)
        {
            if (specBalance.Trim() == "FULL") valid = true;
            else if (IsNum(specBalance)) valid = Num(specBalance) > 0;
        }
        if (afterLimitExInd == "Y") valid = true;
        return valid ?? false;
    }
}
