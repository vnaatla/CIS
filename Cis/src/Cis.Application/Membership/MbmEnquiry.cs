using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Domain.Errors;
using static Cis.Application.Membership.MbmEnquiryRules;

namespace Cis.Application.Membership;

/// <summary>Plan look-ups of the member (cis.MbmEnq_Plan).</summary>
/// <param name="AnnualFound">An AnnualLimit row exists for plan + insured type.</param>
/// <param name="SuppLimitStatus">Its SuppLimitStatus.</param>
/// <param name="PDentalAnnLimit">Its PDentalAnnLimit.</param>
/// <param name="PSpecAnnLimit">Its PSpecAnnLimit (the optical limit).</param>
/// <param name="MaternityFound">A maternity plan's AnnualLimit row exists.</param>
/// <param name="MaternityLimit">Its AnnMaternityLmt.</param>
/// <param name="SpecFound">A SpecialistLimit row exists.</param>
/// <param name="SpecStatus">Its SuppLimitStatus.</param>
/// <param name="MaternityInd">CISDB PLN.PLNMaternityInd.</param>
/// <param name="DentalInd">CISDB PLN.PDentalStatus.</param>
/// <param name="Product">"ProductName(PLNDescription)" of HISDB MBM + PLN, or "".</param>
/// <param name="FlexCount">FlexCare plan rows covering the expiry date.</param>
public record EnqPlan(bool AnnualFound, string? SuppLimitStatus, decimal? PDentalAnnLimit, decimal? PSpecAnnLimit, bool MaternityFound, decimal? MaternityLimit,
    bool SpecFound, string? SpecStatus, string MaternityInd, string DentalInd, string Product, int FlexCount);

/// <summary>GetSpecLimit inputs (cis.MbmEnq_SpecLimit).</summary>
/// <param name="Indicators">PLN.PSpecLimit of plan + payor (in row order).</param>
/// <param name="Found">A SpecialistLimit row exists.</param>
/// <param name="PrinLimit">PrinLimit.</param>
/// <param name="Status">SuppLimitStatus.</param>
public record EnqSpecLimit(List<string?> Indicators, bool Found, decimal? PrinLimit, string? Status);

/// <summary>A covered person in the Supplementary list.</summary>
/// <param name="CoverId">Cover id.</param>
/// <param name="Name">Name.</param>
/// <param name="Ic">IC.</param>
/// <param name="Relationship">REL code.</param>
/// <param name="Dob">Date of birth.</param>
/// <param name="Plan">Cover plan (the principal's when empty).</param>
/// <param name="Status">MBMCCLNStatus.</param>
/// <param name="CardType">Card type.</param>
/// <param name="Premium">Premium.</param>
/// <param name="Mco">MCO fee.</param>
/// <param name="Exclusion">MBMCCLNWarranty.</param>
/// <param name="Remarks">MBMCCLNRemarks.</param>
public record EnqCover(string CoverId, string Name, string Ic, string Relationship, DateTime? Dob, string Plan, string Status, string CardType, decimal Premium,
    decimal Mco, string Exclusion, string Remarks);

/// <summary>The enquiry of one member (cmdShow_Click → FindMBM → showMBM).</summary>
/// <param name="MbmNumber">Member number.</param>
/// <param name="Status">MBMCLNStatus.</param>
/// <param name="ExpiryStatus">MBMCLNExpiryStatus.</param>
/// <param name="DataStatus">MBMCLNDataStatus.</param>
/// <param name="PayCode">Payor.</param>
/// <param name="Name">Name.</param>
/// <param name="PolicyNo">Policy number.</param>
/// <param name="Ic">IC / BC / passport.</param>
/// <param name="InsType">Insured type.</param>
/// <param name="Plan">Plan.</param>
/// <param name="Address">Address lines 1–3.</param>
/// <param name="Country">CTYCode.</param>
/// <param name="PostCode">Post code.</param>
/// <param name="State">STACode.</param>
/// <param name="Subsidiary">Subsidiary.</param>
/// <param name="Department">Department.</param>
/// <param name="GroupCompany">Group company.</param>
/// <param name="Position">Position.</param>
/// <param name="HisRemarks">MBMRemarks.</param>
/// <param name="SpecialRemark">PLNSpecialRemark (shown as a message after Show).</param>
/// <param name="Dob">Date of birth.</param>
/// <param name="AdultChild">A / C.</param>
/// <param name="EffDate">Policy effective date.</param>
/// <param name="ExpDate">Policy expiry date.</param>
/// <param name="AgeBand">Age band code.</param>
/// <param name="CardType">Card type.</param>
/// <param name="Exclusion">MBMCLNWarranty.</param>
/// <param name="CisRemarks">MBMCLNRemarks.</param>
/// <param name="InvoiceDate">MBMCInvoiceDate.</param>
/// <param name="InvoiceNo">MBMCInvoiceNo.</param>
/// <param name="Premium">Premium.</param>
/// <param name="McoFee">MCO fee.</param>
/// <param name="Product">Product caption.</param>
/// <param name="LimitStatus">AnnualLimit SuppLimitStatus ("" when none).</param>
/// <param name="LimitStatusText">Its text.</param>
/// <param name="TopupAmt">Top-up label.</param>
/// <param name="SpTopupAmt">Specialist top-up label.</param>
/// <param name="FfsAmt">FFS label.</param>
/// <param name="AccumAmt">Accumulated (utilised) label.</param>
/// <param name="AnnualLimit">Annual limit label.</param>
/// <param name="BalanceLimit">Balance label.</param>
/// <param name="VisitLimit">Visit limit label.</param>
/// <param name="Gp">GP column of the limits grid.</param>
/// <param name="Specialist">SP column.</param>
/// <param name="Maternity">Maternity column.</param>
/// <param name="Dental">Dental column.</param>
/// <param name="Optical">Optical column.</param>
/// <param name="MepLimit">MBMMEPLmt.</param>
/// <param name="DentalLimit">MBMDentalLmt.</param>
/// <param name="SpecLimit">MBMSpecLmt.</param>
/// <param name="DentalExtLimit">MBMDentalExtLmt.</param>
/// <param name="FirstSpecLimit">MBMFirstSpecLmt.</param>
/// <param name="Covers">Covered persons (F / H).</param>
/// <param name="FlexCare">The FlexCare / Family Benefits tabs apply.</param>
/// <param name="Warnings">The desktop's look-up warnings.</param>
public record EnqMember(string MbmNumber, string Status, string ExpiryStatus, string DataStatus, string PayCode, string Name, string PolicyNo, string Ic,
    string InsType, string Plan, List<string> Address, string Country, string PostCode, string State, string Subsidiary, string Department, string GroupCompany,
    string Position, string HisRemarks, string SpecialRemark, DateTime? Dob, string AdultChild, DateTime? EffDate, DateTime? ExpDate, string AgeBand, string CardType,
    string Exclusion, string CisRemarks, DateTime? InvoiceDate, string InvoiceNo, decimal Premium, decimal McoFee, string Product, string LimitStatus,
    string LimitStatusText, string TopupAmt, string SpTopupAmt, string FfsAmt, string AccumAmt, string AnnualLimit, string BalanceLimit, string VisitLimit,
    EnqBenefit Gp, EnqBenefit Specialist, EnqBenefit Maternity, EnqBenefit Dental, EnqBenefit Optical, decimal MepLimit, decimal DentalLimit, decimal SpecLimit,
    decimal DentalExtLimit, decimal FirstSpecLimit, List<EnqCover> Covers, bool FlexCare, List<string> Warnings);

/// <summary>A case of the Case tab.</summary>
/// <param name="CaseNo">Case number.</param>
/// <param name="MbmNumber">Member number.</param>
/// <param name="Name">Patient name.</param>
/// <param name="Relationship">P for the principal, else the relationship code.</param>
/// <param name="VisitDate">Visit date.</param>
/// <param name="Claimability">Claimability.</param>
/// <param name="Reason">Reason.</param>
/// <param name="Status">Case status.</param>
/// <param name="Bill">CLMTotalActual.</param>
/// <param name="Bordx">CLMBordxAmt.</param>
/// <param name="Ffs">CLMMCOProcFee.</param>
public record EnqCase(string CaseNo, string MbmNumber, string Name, string Relationship, DateTime? VisitDate, string Claimability, string Reason, string Status,
    decimal Bill, decimal Bordx, decimal Ffs);

/// <summary>Case tab: GP and specialist cases with the desktop's totals.</summary>
/// <param name="Gp">GP cases.</param>
/// <param name="GpVisits">Total visits.</param>
/// <param name="GpBill">Total bill.</param>
/// <param name="GpBordx">Total bordereaux (not deleted).</param>
/// <param name="GpFfs">Total FFS (not deleted).</param>
/// <param name="Specialist">Specialist cases.</param>
/// <param name="SpVisits">Specialist visits.</param>
/// <param name="SpBill">Specialist bill.</param>
/// <param name="SpBordx">Specialist "bordereaux" (the desktop adds the bill amount).</param>
public record EnqCases(List<EnqCase> Gp, int GpVisits, decimal GpBill, decimal GpBordx, decimal GpFfs, List<EnqCase> Specialist, int SpVisits, decimal SpBill,
    decimal SpBordx);

/// <summary>Data access of the MBM Enquiry.</summary>
public interface IMbmEnquiryRepository
{
    /// <summary>SearchCISMemberProc_test rows.</summary>
    Task<List<Dictionary<string, object?>>> MemberAsync(IDbSession session, string mbmNumber);

    /// <summary>Plan look-ups.</summary>
    Task<EnqPlan> PlanAsync(IDbSession session, string mbmNumber, string plan, string insType, DateTime? expDate);

    /// <summary>Top-up totals (null = no row).</summary>
    Task<(decimal Topup, decimal SpTopup)?> TopupAsync(IDbSession session, string mbmNumber, string? covId);

    /// <summary>Tbl_ASOMemberDetails (last row: ASO limit, OP limit, combined).</summary>
    Task<(decimal Aso, decimal Op, string Combined)?> AsoAsync(IDbSession session, string mbmNumber, string covId);

    /// <summary>GetSpecLimit inputs.</summary>
    Task<EnqSpecLimit> SpecLimitAsync(IDbSession session, string plan, string payCode, string insType);

    /// <summary>The member's claims (CIS_FFS2018_test, all covers).</summary>
    Task<List<EnqClaimRow>> ClaimsAsync(IDbSession session, string mbmNumber);

    /// <summary>tmpDlimit_local_test1.</summary>
    Task<List<EnqDentalRow>> DentalAsync(IDbSession session, string mbmNumber, string covId);

    /// <summary>tmpDentallimit_OneLimit.</summary>
    Task<List<EnqDentalRow>> DentalOneLimitAsync(IDbSession session, string mbmNumber);

    /// <summary>GP (false) or specialist (true) cases.</summary>
    Task<List<Dictionary<string, object?>>> CasesAsync(IDbSession session, string mbmNumber, bool specialist);
}

/// <summary>Membership > MBM Enquiry (FrmMBMCLNEnq).</summary>
public interface IMbmEnquiryService
{
    /// <summary>cmdShow_Click.</summary>
    Task<EnqMember> FindAsync(string? member);

    /// <summary>Case tab (ListCases + ListSpecialistCases).</summary>
    Task<EnqCases> CasesAsync(string? member);

    /// <summary>Benefit Schedule tab (LoadDoc): .jpg / .pdf files of the group company's policy wording folder.</summary>
    List<string>? PolicyFiles(string? groupCompany);
}

/// <inheritdoc />
public sealed class MbmEnquiryService(IDbSessions sessions, IMbmEnquiryRepository repository, IMbmAdjustmentRepository adjustment, IGeneralExclusionsService exclusions)
    : IMbmEnquiryService
{
    /// <summary>Desktop texts.</summary>
    public const string EnterMember = "Please Enter Membership No. Atleast first 8 characters required.", NotFound = "No record found!",
        SpecWarning = "Specialist Limit Information Is Not Correct! \nPlease check you Maintenance Table";

    /// <summary>LoadDoc lists these extensions only (case as written).</summary>
    public static readonly string[] DocExtensions = [".jpg", ".JPG", ".pdf", ".PDF"];

    /// <inheritdoc />
    public async Task<EnqMember> FindAsync(string? member)
    {
        var number = Number(member);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.MemberAsync(s, number);
        if (rows.Count == 0) throw new NotFoundException(NotFound);
        var p = rows[0];
        var warnings = new List<string>();
        string ins = T(p, "INSType"), plan = T(p, "PLNType"), pay = T(p, "PAYCode");
        DateTime? eff = p.GetValueOrDefault("MBMPolEffDate") as DateTime?, exp = p.GetValueOrDefault("MBMPolExpDate") as DateTime?,
            dob = p.GetValueOrDefault("MBMDOB") as DateTime?;
        var pl = await repository.PlanAsync(s, number, plan, ins, exp);
        var limitStatus = pl.AnnualFound ? Blank(pl.SuppLimitStatus, "A") : "";
        var flags = new EnqFlags(T(p, "PDentalStatus"), T(p, "SDentalStatus"), T(p, "IsGPDentalShared"), T(p, "IsDentalOpticalShared"), T(p, "PSpecStatus"),
            T(p, "SSpecStatus"), Pdl(pl.AnnualFound, pl.PDentalAnnLimit));
        string opticalLimit = pl.AnnualFound ? Cstr(pl.PSpecAnnLimit, "0.00") : "";

        string topup = NA, spTopup = NA;
        if (limitStatus is "A" or "B" or "C" or "D" or "E" or "F")
        {
            var t = await repository.TopupAsync(s, number, limitStatus == "A" ? null : "00");
            (topup, spTopup) = t is { } v ? (Money(v.Topup), Money(v.SpTopup)) : (NA, "0");
        }

        var (annual, visit, specLimit) = await CallLimit(s, p, number, ins, plan, pay, eff, limitStatus, warnings);
        var claims = await repository.ClaimsAsync(s, number);
        var shared = p.GetValueOrDefault("IsPrinDepGPShared") is true or 1;
        bool maternityPlan = pl.MaternityInd == "Y", dentalPlan = pl.DentalInd == "Y";
        decimal gpFfs = 0, utilised = 0;
        string ffsLabel, accum, specAccum, specBal;
        if (specLimit == NA)
        {
            foreach (var r in Filter(claims, Mode(limitStatus, shared, false)))
            {
                gpFfs += r.Ffs ?? 0;
                utilised += GpAdd(flags, maternityPlan, r) ?? 0;
            }
            ffsLabel = gpFfs.ToString("#,###.00", CultureInfo.InvariantCulture);
            accum = Accum(utilised);
            (specAccum, specBal) = (NA, NA);
        }
        else
        {
            foreach (var r in Filter(claims, Mode(limitStatus, shared, true)))
            {
                gpFfs += r.Ffs ?? 0;
                utilised += GpAddWithSpec(maternityPlan, dentalPlan, r) ?? 0;
            }
            decimal specFfs = 0, specAcc = 0;
            var specStatus = pl.SpecFound ? Blank(pl.SpecStatus, "A") : "";
            foreach (var r in Filter(claims, specStatus switch { "A" => 0, "B" or "C" or "" or "D" or "E" => 1, _ => (int?)null })
                         .Where(r => r.ClnType?.Trim() == "S"))
            {
                specFfs += r.Ffs ?? 0;
                specAcc += (r.ClmBordxAmt ?? 0) - (r.ClmOthersNoDed ?? 0) - (r.ClmMaternity ?? 0);
            }
            ffsLabel = Money(gpFfs + specFfs);
            accum = Accum(utilised);
            if (IsNum(specLimit))
            {
                specLimit = Money(specLimit);
                (specAccum, specBal) = (Accum(specAcc), Money(Num(specLimit) + (IsNum(spTopup) ? Num(spTopup) : 0) - specAcc));
            }
            else if (specLimit == "FULL") (specAccum, specBal) = (specAcc.ToString("#,###.00", CultureInfo.InvariantCulture), "FULL");
            else (specLimit, specAccum, specBal, spTopup) = (NA, NA, NA, NA);
        }

        var annLabel = limitStatus == "E" ? "FULL" : annual;
        if (IsNum(annual)) annLabel = Money(annLabel);
        var balance = IsNum(annLabel) ? Money(Num(annLabel) + (IsNum(topup) ? Num(topup) : 0) - utilised) : annLabel == "FULL" ? "FULL" : NA;
        var visitLabel = IsNum(visit) ? Money(N(p, "MBMCLNVisitLmt")) : visit;

        var matLimit = pl.MaternityFound ? Cstr(pl.MaternityLimit, "") : "0";
        var matAcc = claims.Sum(r => r.ClmMaternity ?? 0);
        var maternity = IsNum(matLimit) ? new EnqBenefit(Money(matLimit), Money(matAcc), Money(Num(matLimit) - matAcc)) : new EnqBenefit("0.00", "0.00", "0.00");

        var optical = Optical(claims);
        var source = DentalSource(flags);
        var asoType = T(p, "ASOType");
        var pri = source == 1 ? PrincipalDental(flags, await repository.DentalAsync(s, number, "00")) : (0m, 0m);
        var one = source == 2 ? OneLimitDental(flags, asoType, await repository.DentalOneLimitAsync(s, number)) : (0m, 0m, 0m, 0m);
        var dental = DentalLabels(flags, asoType, pri, one, optical);
        var vision = Vision(flags, optical, opticalLimit);

        var covers = ins is "F" or "H"
            ? rows.Where(r => T(r, "MBMCCoverID") is { Length: > 0 } id && id != "00")
                .Select(c => new EnqCover(T(c, "MBMCCoverID"), T(c, "MBMCName"), T(c, "MBMCICBCPPNo"), T(c, "MBMCRELCode"), c.GetValueOrDefault("MBMCDOB") as DateTime?,
                    T(c, "MBMCCLNPLNCode") is { Length: > 0 } cp ? cp : plan, T(c, "MBMCCLNStatus"), T(c, "MBMCCardType"), N(c, "MBMCCLNPrem"), N(c, "MBMCCLNMCO"),
                    T(c, "MBMCCLNWarranty"), T(c, "MBMCCLNRemarks"))).ToList()
            : [];
        var ageBand = eff is null || dob is null ? "" : await AgeBand(s, eff.Value, dob.Value);

        return new EnqMember(number, T(p, "MBMCLNStatus"), T(p, "MBMCLNExpiryStatus"), T(p, "MBMCLNDataStatus"), pay, T(p, "MBMName"), T(p, "MBMPolicyNo"),
            T(p, "MBMIcBcPp"), ins, plan, [T(p, "MBMAddress1"), T(p, "MBMAddress2"), T(p, "MBMAddress3")], T(p, "CTYCode"), T(p, "MBMPostCode"), T(p, "STACode"),
            T(p, "Subsidiary"), T(p, "MBMDepartment"), T(p, "MBMGroupCompany"), T(p, "MBMPosition"), T(p, "MBMRemarks"), T(p, "PLNSpecialRemark"), dob,
            T(p, "MBMAdultChild"), eff, exp, ageBand, T(p, "MBMCardType"), T(p, "MBMCLNWarranty"), T(p, "MBMCLNRemarks"),
            p.GetValueOrDefault("MBMCInvoiceDate") as DateTime?, T(p, "MBMCInvoiceNo"), N(p, "MBMCLNPremium"), N(p, "MBMCLNMCOFee"), pl.Product, limitStatus,
            limitStatus == "" ? "" : MbmAdjustmentService.LimitStatusText(limitStatus), topup, spTopup, ffsLabel, accum, annLabel, balance, visitLabel,
            new EnqBenefit(Money(annLabel), Money(accum), Money(balance)), new EnqBenefit(Money(specLimit), Money(specAccum), Money(specBal)), maternity,
            new EnqBenefit(Money(dental.Limit), Money(dental.Utilised), Money(dental.Balance)),
            new EnqBenefit(Money(vision.Limit), Money(vision.Utilised), Money(vision.Balance)), N(p, "MBMMEPLmt"), N(p, "MBMDentalLmt"), N(p, "MBMSpecLmt"),
            N(p, "MBMDentalExtLmt"), N(p, "MBMFirstSpecLmt"), covers, exp is not null && pl.FlexCount == 1, warnings);
    }

    /// <inheritdoc />
    public async Task<EnqCases> CasesAsync(string? member)
    {
        var number = Number(member);
        await using var s = await sessions.OpenAsync();
        var gp = (await repository.CasesAsync(s, number, false)).Select(Case).ToList();
        var sp = (await repository.CasesAsync(s, number, true)).Select(Case).ToList();
        var live = gp.Where(c => c.Status != "D").ToList();
        return new EnqCases(gp, gp.Count, gp.Sum(c => c.Bill), live.Sum(c => c.Bordx), live.Sum(c => c.Ffs), sp, sp.Count, sp.Sum(c => c.Bill), sp.Sum(c => c.Bill));
    }

    /// <inheritdoc />
    public List<string>? PolicyFiles(string? groupCompany) =>
        string.IsNullOrWhiteSpace(groupCompany) ? [] : exclusions.PolicyFiles(groupCompany)?.Where(f => DocExtensions.Contains(Path.GetExtension(f), StringComparer.Ordinal)).ToList();

    /// <summary>The cover filter of showMBM's utilisation: 0 all covers, 1 cover 00, null = no query (zero).</summary>
    public static int? Mode(string limitStatus, bool shared, bool specBranch) => shared ? 0 : limitStatus switch
    {
        "A" => 0,
        "B" or "C" or "D" or "E" or "" => 1,
        "F" => specBranch ? null : 1,
        _ => null,
    };

    /// <summary>Medix.bas GetSpecLimit ("FULL", "N/A" or the amount) for the principal.</summary>
    public static string SpecLimit(EnqSpecLimit v, List<string> warnings)
    {
        var ind = "N";
        foreach (var i in v.Indicators) ind = i?.Trim() is "F" or "Y" ? i.Trim() : "N";
        if (ind == "F") return "FULL";
        if (ind != "Y") return NA;
        if (!v.Found) return "0";
        if (string.IsNullOrEmpty(v.Status))
        {
            warnings.Add(SpecWarning);
            return "FULL";
        }
        return v.Status == "D" ? "FULL" : Cstr(v.PrinLimit ?? 0, "0");
    }

    private async Task<(string Annual, string Visit, string Spec)> CallLimit(IDbSession s, Dictionary<string, object?> p, string number, string ins, string plan,
        string pay, DateTime? eff, string limitStatus, List<string> warnings)
    {
        string annual = "", visit = "";
        LimitLookup? limits = eff is null ? null : await adjustment.LimitsAsync(s, "P", ins, plan, eff.Value, pay, limitStatus);
        if (T(p, "ASOInd") == "Y")
        {
            var aso = await repository.AsoAsync(s, number, "00");
            annual = Cstr(aso is { Combined: "Y" } a ? a.Aso : aso?.Op ?? 0, "0");
        }
        else if (limits is not null)
        {
            if (!limits.AnnualOk) warnings.Add(MbmAdjustmentService.NoAnnual);
            annual = limits.AnnualFull ? "FULL" : Money(limits.AnnualLimit);
        }
        if (limits is not null) visit = limits.VisitFull ? "FULL" : Money(limits.VisitLimit);
        var spec = SpecLimit(await repository.SpecLimitAsync(s, plan, pay, ins), warnings);
        return (annual, visit, IsNum(spec) ? Money(spec) : spec);
    }

    private async Task<string> AgeBand(IDbSession s, DateTime eff, DateTime dob)
    {
        var age = MemberRegistrationService.GetAge(eff, dob);
        var code = await adjustment.AgeCodeAsync(s, int.Parse(age[..^1], CultureInfo.InvariantCulture), age.EndsWith('D'));
        return code.Length >= 2 ? code[..2] : code.PadLeft(2, '0');
    }

    private static IEnumerable<EnqClaimRow> Filter(List<EnqClaimRow> rows, int? mode) => mode switch
    {
        0 => rows,
        1 => rows.Where(r => r.CoverId?.Trim() == "00"),
        _ => [],
    };

    private static EnqCase Case(Dictionary<string, object?> r)
    {
        var principal = T(r, "CoverID") == "00";
        return new EnqCase(T(r, "CASNumber"), T(r, "MBMNumber"), principal ? T(r, "MBMName") : T(r, "MBMCName"), principal ? "P" : T(r, "MBMCRELCode"),
            r.GetValueOrDefault("DateVisit") as DateTime?, T(r, "Claimability"), T(r, "ReasonDesc"), T(r, "CASStatus"), N(r, "CLMTotalActual"), N(r, "CLMBordxAmt"),
            N(r, "CLMMCOProcFee"));
    }

    private static string Accum(decimal v) => v.ToString("#,###.000", CultureInfo.InvariantCulture);

    private static string Blank(string? v, string empty) => string.IsNullOrWhiteSpace(v) ? empty : v.Trim();

    private static string Cstr(decimal? v, string empty) => v is null ? empty : v.Value.ToString("0.############", CultureInfo.InvariantCulture);

    private static string Number(string? member) =>
        string.IsNullOrWhiteSpace(member) || member.Trim().Length < 8 ? throw new BusinessRuleException(EnterMember) : member.Trim().ToUpperInvariant();

    private static string T(Dictionary<string, object?> row, string key) =>
        row.TryGetValue(key, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(Dictionary<string, object?> row, string key) => row.GetValueOrDefault(key) switch
    {
        decimal d => d,
        double d => (decimal)d,
        int i => i,
        _ => 0,
    };
}
