using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Domain.Errors;
using static Cis.Application.Membership.MbmEnquiryRules;

namespace Cis.Application.Membership;

/// <summary>Plan look-ups of the member (cis.MbmEnq_Plan).</summary>
/// <param name="AnnualFound">An AnnualLimit row exists for plan + insured type.</param>
/// <param name="SuppLimitStatus">Its SuppLimitStatus.</param>
/// <param name="PDentalAnnLimit">Its PDentalAnnLimit.</param>
/// <param name="SDentalAnnLimit">Its SDentalAnnLimit.</param>
/// <param name="PSpecAnnLimit">Its PSpecAnnLimit (the optical limit).</param>
/// <param name="MaternityFound">A maternity plan's AnnualLimit row exists.</param>
/// <param name="MaternityLimit">Its AnnMaternityLmt.</param>
/// <param name="SpecFound">A SpecialistLimit row exists.</param>
/// <param name="SpecStatus">Its SuppLimitStatus.</param>
/// <param name="MaternityInd">CISDB PLN.PLNMaternityInd.</param>
/// <param name="DentalInd">CISDB PLN.PDentalStatus.</param>
/// <param name="Product">"ProductName(PLNDescription)" of HISDB MBM + PLN, or "".</param>
/// <param name="FlexCount">FlexCare plan rows covering the expiry date.</param>
public record EnqPlan(bool AnnualFound, string? SuppLimitStatus, decimal? PDentalAnnLimit, decimal? SDentalAnnLimit, decimal? PSpecAnnLimit, bool MaternityFound, decimal? MaternityLimit,
    bool SpecFound, string? SpecStatus, string MaternityInd, string DentalInd, string Product, int FlexCount);

/// <summary>GetSpecLimit inputs (cis.MbmEnq_SpecLimit).</summary>
/// <param name="Indicators">PLN.PSpecLimit of plan + payor (in row order).</param>
/// <param name="Found">A SpecialistLimit row exists.</param>
/// <param name="PrinLimit">PrinLimit.</param>
/// <param name="Status">SuppLimitStatus.</param>
/// <param name="SuppLimit">SuppLimit (supplementary members when the status is not A).</param>
public record EnqSpecLimit(List<string?> Indicators, bool Found, decimal? PrinLimit, string? Status, decimal? SuppLimit = null);

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
/// <param name="PremReduction">Premium reduction (cancelled membership, MBMCISRefund).</param>
/// <param name="PremEffective">Effective premium.</param>
/// <param name="McoRefund">MCO refund.</param>
/// <param name="McoEffective">Effective MCO.</param>
/// <param name="Gp">GP limit / utilised / balance (S16–S18).</param>
/// <param name="Specialist">Specialist (S27–S29).</param>
/// <param name="Dental">Dental (S32–S34).</param>
/// <param name="Optical">Optical (S36).</param>
/// <param name="VisitLimit">Visit limit.</param>
/// <param name="MepLimit">MEP limit.</param>
/// <param name="DentalLimit">Dental limit.</param>
/// <param name="DentalExtLimit">Dental extraction limit.</param>
/// <param name="SpecLimit">Spectacles limit.</param>
/// <param name="FirstSpecLimit">First spectacles limit.</param>
/// <param name="AgeBand">Age band.</param>
/// <param name="AdultChild">A / C.</param>
/// <param name="EffDate">Cover effective date.</param>
/// <param name="ExpDate">Cover expiry date.</param>
/// <param name="Ffs">lstCovAdd_Click FFS label.</param>
/// <param name="TopupAmt">lstCovAdd_Click top-up label.</param>
/// <param name="SpTopupAmt">Specialist top-up label.</param>
public record EnqCover(string CoverId, string Name, string Ic, string Relationship, DateTime? Dob, string Plan, string Status, string CardType, decimal Premium,
    decimal Mco, string Exclusion, string Remarks, decimal? PremReduction, decimal? PremEffective, decimal? McoRefund, decimal? McoEffective, EnqBenefit Gp,
    EnqBenefit Specialist, EnqBenefit Dental, EnqBenefit Optical, string VisitLimit, decimal MepLimit, decimal DentalLimit, decimal DentalExtLimit, decimal SpecLimit,
    decimal FirstSpecLimit, string AgeBand, string AdultChild, DateTime? EffDate, DateTime? ExpDate, string Ffs, string TopupAmt, string SpTopupAmt);

/// <summary>showMBMCov, limit status F: PLN + AnnualLimit of the plan.</summary>
/// <param name="AnnLmtIndSupp">ANNLmtIndSupp.</param>
/// <param name="SuppLimit">SuppLimit.</param>
/// <param name="ChildDepLimit">ChildDepLimit.</param>
/// <param name="SSpecLimit">SSpecLimit.</param>
/// <param name="VisLmtInd">VISLmtInd.</param>
public record EnqFLimit(string? AnnLmtIndSupp, decimal? SuppLimit, decimal? ChildDepLimit, string? SSpecLimit, string? VisLmtInd);

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
/// <param name="PrincipalNote">MessageDisplay note of the principal (UB category plans; empty otherwise).</param>
/// <param name="DependantNote">MessageDisplay note of the dependants.</param>
public record EnqMember(string MbmNumber, string Status, string ExpiryStatus, string DataStatus, string PayCode, string Name, string PolicyNo, string Ic,
    string InsType, string Plan, List<string> Address, string Country, string PostCode, string State, string Subsidiary, string Department, string GroupCompany,
    string Position, string HisRemarks, string SpecialRemark, DateTime? Dob, string AdultChild, DateTime? EffDate, DateTime? ExpDate, string AgeBand, string CardType,
    string Exclusion, string CisRemarks, DateTime? InvoiceDate, string InvoiceNo, decimal Premium, decimal McoFee, string Product, string LimitStatus,
    string LimitStatusText, string TopupAmt, string SpTopupAmt, string FfsAmt, string AccumAmt, string AnnualLimit, string BalanceLimit, string VisitLimit,
    EnqBenefit Gp, EnqBenefit Specialist, EnqBenefit Maternity, EnqBenefit Dental, EnqBenefit Optical, decimal MepLimit, decimal DentalLimit, decimal SpecLimit,
    decimal DentalExtLimit, decimal FirstSpecLimit, List<EnqCover> Covers, bool FlexCare, List<string> Warnings, string PrincipalNote = "", string DependantNote = "");

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

/// <summary>A row of the FlexCare grids: item with capped limit / utilised / balance (blank when the grid has no such column).</summary>
/// <param name="Item">Caption.</param>
/// <param name="Capped">Capped limit.</param>
/// <param name="Utilised">Utilised.</param>
/// <param name="Balance">Balance.</param>
/// <param name="Highlight">Yellow cell in the desktop.</param>
public record EnqFlexRow(string Item, string Capped, string Utilised, string Balance, bool Highlight = false);

/// <summary>FlexCare (cover 00) or Family Benefits (other covers) of a UOB FlexCare member (ShowFlexiMBM / PriFlexiCareMember).</summary>
/// <param name="Found">spGet_UOB_Flex_New_Cal returned rows (else the tabs are hidden).</param>
/// <param name="CoverId">Cover shown.</param>
/// <param name="CappedItems">FlexCareGridLimits.</param>
/// <param name="Items">FlexCareGridUtility.</param>
/// <param name="Totals">FlexCareGridTotals.</param>
/// <param name="Family">FamilyBenefitsGrid (shown for cover 01 only).</param>
/// <param name="Outpatient">OpUtilisedGrid.</param>
/// <param name="OtherItems">OtherItemUtilisedGrid.</param>
public record EnqFlex(bool Found, string CoverId, List<EnqFlexRow> CappedItems, List<EnqFlexRow> Items, List<EnqFlexRow> Totals, List<EnqFlexRow> Family,
    List<EnqFlexRow> Outpatient, List<EnqFlexRow> OtherItems);

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

    /// <summary>AnnualLimit / SpecialistLimit statuses of a plan on a date (rows latest first).</summary>
    Task<(List<string?> Annual, List<string?> Specialist)> CoverStatusAsync(IDbSession session, string plan, string insType, DateTime effDate);

    /// <summary>PLN + AnnualLimit of a plan (limit status F).</summary>
    Task<EnqFLimit?> FLimitAsync(IDbSession session, string plan);

    /// <summary>CIS_FFS2013 FFS rows: covers other than 00 (2) or one cover (3).</summary>
    Task<List<decimal?>> CoverFfsAsync(IDbSession session, string mbmNumber, int coverMode, string covId);

    /// <summary>cis.MbmEnq_Flex (spGet_UOB_Flex_New_Cal) rows of the member's family for a policy year.</summary>
    Task<List<Dictionary<string, object?>>> FlexAsync(IDbSession session, string mbmNumber, int policyYear);

    /// <summary>CategoryAmountCal (last row: limit, utilised, balance) and Dental_Utilization_Limit_Sum (principal / dependant sums).</summary>
    Task<((decimal Limit, decimal Util, decimal Balance)? HomeOffice, decimal Principal, decimal Dependant)> UbNotesAsync(IDbSession session, string mbmNumber);

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

    /// <summary>FlexCare / Family Benefits tab of a cover (ShowFlexiMBM).</summary>
    Task<EnqFlex> FlexAsync(string? member, string? coverId);

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

        var (annual, visit, specLimit) = await CallLimit(s, p, number, "00", "P", ins, plan, pay, eff, limitStatus, warnings);
        var specLimitRaw = specLimit;
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

        var covers = ins is "F" or "H" ? await Covers(s, rows, p, number, flags, shared, limitStatus, annual, visit, specLimitRaw, claims, pl, warnings) : [];
        var ageBand = eff is null || dob is null ? "" : await AgeBand(s, eff.Value, dob.Value);

        string pNote = "", dNote = "";
        if (number.StartsWith("UB", StringComparison.Ordinal) && N(p, "PlncategoryInd") != 0)
        {
            var (home, priUtil, suppUtil) = await repository.UbNotesAsync(s, number);
            (pNote, dNote) = UbNotes(flags, pl.AnnualFound ? Blank(pl.SuppLimitStatus, "A") : "N", pl.AnnualFound ? pl.PDentalAnnLimit ?? 0 : 0,
                pl.AnnualFound ? pl.SDentalAnnLimit ?? 0 : 0, priUtil, suppUtil, true, home ?? (0, 0, 0));
        }

        return new EnqMember(number, T(p, "MBMCLNStatus"), T(p, "MBMCLNExpiryStatus"), T(p, "MBMCLNDataStatus"), pay, T(p, "MBMName"), T(p, "MBMPolicyNo"),
            T(p, "MBMIcBcPp"), ins, plan, [T(p, "MBMAddress1"), T(p, "MBMAddress2"), T(p, "MBMAddress3")], T(p, "CTYCode"), T(p, "MBMPostCode"), T(p, "STACode"),
            T(p, "Subsidiary"), T(p, "MBMDepartment"), T(p, "MBMGroupCompany"), T(p, "MBMPosition"), T(p, "MBMRemarks"), T(p, "PLNSpecialRemark"), dob,
            T(p, "MBMAdultChild"), eff, exp, ageBand, T(p, "MBMCardType"), T(p, "MBMCLNWarranty"), T(p, "MBMCLNRemarks"),
            p.GetValueOrDefault("MBMCInvoiceDate") as DateTime?, T(p, "MBMCInvoiceNo"), N(p, "MBMCLNPremium"), N(p, "MBMCLNMCOFee"), pl.Product, limitStatus,
            limitStatus == "" ? "" : MbmAdjustmentService.LimitStatusText(limitStatus), topup, spTopup, ffsLabel, accum, annLabel, balance, visitLabel,
            new EnqBenefit(Money(annLabel), Money(accum), Money(balance)), new EnqBenefit(Money(specLimit), Money(specAccum), Money(specBal)), maternity,
            new EnqBenefit(Money(dental.Limit), Money(dental.Utilised), Money(dental.Balance)),
            new EnqBenefit(Money(vision.Limit), Money(vision.Utilised), Money(vision.Balance)), N(p, "MBMMEPLmt"), N(p, "MBMDentalLmt"), N(p, "MBMSpecLmt"),
            N(p, "MBMDentalExtLmt"), N(p, "MBMFirstSpecLmt"), covers, exp is not null && pl.FlexCount == 1, warnings, pNote, dNote);
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
    public async Task<EnqFlex> FlexAsync(string? member, string? coverId)
    {
        var number = Number(member);
        var cov = string.IsNullOrWhiteSpace(coverId) ? "00" : coverId.Trim();
        await using var s = await sessions.OpenAsync();
        var p = (await repository.MemberAsync(s, number)).FirstOrDefault() ?? throw new NotFoundException(NotFound);
        var exp = p.GetValueOrDefault("MBMPolExpDate") as DateTime?;
        // cmdShow_Click / lstCovAdd_Click call ShowFlexiMBM only for FlexCare plans (UbCount = 1); without a family the procedure is unfiltered.
        var flexPlan = exp is not null && (await repository.PlanAsync(s, number, T(p, "PLNType"), T(p, "INSType"), exp)).FlexCount == 1;
        var rows = flexPlan ? await repository.FlexAsync(s, number, exp!.Value.Year) : [];
        return Flex(rows, number, cov);
    }

    /// <summary>ShowFlexiMBM + PriFlexiCareMember on the procedure's rows.</summary>
    public static EnqFlex Flex(List<Dictionary<string, object?>> rows, string number, string cov)
    {
        if (rows.Count == 0) return new EnqFlex(false, cov, [], [], [], [], [], []);
        var v = new Dictionary<string, string>();
        string G(string k) => v.TryGetValue(k, out var x) ? x : "0.00";
        foreach (var r in rows.Where(r => T(r, "CoverID") == cov && T(r, "MemberNo") == number))
        {
            if (cov == "00")
                foreach (var col in new[] { "619", "300", "618", "266", "617Cap", "617", "617Bal", "466Cap", "466", "466Bal", "703Cap", "703", "703Bal", "697", "696", "654" })
                    v[col] = FlexValue(r, col);
            foreach (var col in new[] { "615", "303", "468", "616", "305", "135", "614", "306" }) v[col] = FlexValue(r, col);
            if (N(r, "MBMCountNo") == 0)
            {
                v["limit"] = FlexValue(r, "Limits");
                v["util"] = FlexValue(r, "SumFlexAmount");
                v["sumop"] = FlexValue(r, "SumOP");
                v["cf"] = FlexValue(r, "PriFlexBalCF");
                v["leaves"] = FlexValue(r, "LeavesSoldAmount");
                v["reward"] = FlexValue(r, "Reward");
                v["current"] = FlexValue(r, "PriFlexAvailLimit");
                v["balance"] = Money(Num(v["current"]) - Num(v["util"]));
            }
            else
            {
                v["limit"] = FlexValue(r, "Limits");
                v["deputil"] = FlexValue(r, "DepUtil");
                v["depbal"] = FlexValue(r, "DepBal");
                v["sumop"] = Money(Num(FlexValue(r, "SumOP")) - Num(FlexValue(r, "SumFlexAmount")));
            }
        }
        EnqFlexRow U(string item, string col) => new(item, "", G(col), "");
        var others = new[]
        {
            ("Chiropractic", "615"), ("Dental", "303"), ("Health Screening", "306"), ("Insurance claim", "468"), ("Mental health treatment/counselling", "616"),
            ("Optical", "305"), ("Traditional medical treatment", "135"), ("Vaccination", "614"),
        };
        return new EnqFlex(true, cov,
            [
                new("Sport equipment", G("617Cap"), G("617"), G("617Bal")), new("Vitamins/health supplements", G("466Cap"), G("466"), G("466Bal")),
                new("Home office", G("703Cap"), G("703"), G("703Bal")), new("COVID-19 Screening(Flexcare)", "0.00", G("654"), "0.00"),
                new("Medical Devices", "0.00", G("697"), "0.00"), new("Physiotherapy", "0.00", G("696"), "0.00"),
            ],
            [
                U("Car Park", "619"), U("Childcare expenses", "618"), U("Chiropratic", "615"), U("Dental", "303"), U("Gym", "266"), U("Health Screening", "306"),
                U("Maternity", "300"), U("Mental Health treatment/Counselling", "616"), U("Optical", "305"), U("Traditional medical treatment", "135"),
                U("Vacciantion", "614"),
            ],
            [
                new("Total Flex Care Limit", "", G("limit"), ""), new("FlexCARE Balance Carry Forward from Previous Year", "", G("cf"), ""),
                new("Annual Leaves sold conversion", "", G("leaves"), ""), new("Rewards Amount", "", G("reward"), ""),
                new("Total flexCARE for current year", "", G("current"), "", true), new("Total Utilised", "", G("util"), "", true),
                new("Total flexCARE Balance Available", "", G("balance"), "", true),
            ],
            [
                new("Total Family Medical Benefit Limit", "", G("limit"), ""), new("Total Utilised by Family", "", G("deputil"), ""),
                new("Total Family Medical Benefit Balance available", "", G("depbal"), ""),
            ],
            [new("GP", "", G("sumop"), ""), new("SP", "", "0.00", ""), new("Inpatient", "", "-", ""), new("Total", "", G("sumop"), "")],
            [.. others.Select(o => U(o.Item1, o.Item2)), new("Total", "", Money(others.Sum(o => Num(G(o.Item2)))), "")]);
    }

    /// <summary>Format(IIf(Len(x), x, 0), "#,##0.00"): amounts formatted, texts such as "NA" kept.</summary>
    private static string FlexValue(Dictionary<string, object?> r, string col)
    {
        var t = r.GetValueOrDefault(col) is { } o ? Convert.ToString(o, CultureInfo.InvariantCulture)!.Trim() : "";
        return t.Length == 0 ? "0.00" : IsNum(t) ? Money(t) : t;
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

    /// <summary>Medix.bas GetSpecLimit ("FULL", "N/A" or the amount); supplementary members (MBMType S) take SuppLimit unless the status is A.</summary>
    public static string SpecLimit(EnqSpecLimit v, List<string> warnings, bool supplementary = false)
    {
        var ind = "N";
        foreach (var i in v.Indicators) ind = i?.Trim() is "F" or "Y" ? i.Trim() : "N";
        if (ind == "F") return "FULL";
        if (ind != "Y") return NA;
        if (!v.Found) return "0";
        string limit;
        if (string.IsNullOrEmpty(v.Status))
        {
            warnings.Add(SpecWarning);
            limit = "FULL";
        }
        else limit = v.Status == "D" ? "FULL" : Cstr(v.PrinLimit ?? 0, "0");
        return supplementary && v.Status != "A" ? Cstr(v.SuppLimit ?? 0, "0") : limit;
    }

    /// <summary>CallLimit / CallLimitDept: annual (ASO members: Tbl_ASOMemberDetails of the cover), visit and specialist limits.</summary>
    private async Task<(string Annual, string Visit, string Spec)> CallLimit(IDbSession s, Dictionary<string, object?> p, string number, string covId, string mbmType,
        string ins, string plan, string pay, DateTime? eff, string limitStatus, List<string> warnings)
    {
        string annual = "", visit = "";
        LimitLookup? limits = eff is null ? null : await adjustment.LimitsAsync(s, mbmType, ins, plan, eff.Value, pay, limitStatus);
        if (T(p, "ASOInd") == "Y")
        {
            var aso = await repository.AsoAsync(s, number, covId);
            annual = Cstr(aso is { Combined: "Y" } a ? a.Aso : aso?.Op ?? 0, "0");
        }
        else if (limits is not null)
        {
            if (!limits.AnnualOk) warnings.Add(MbmAdjustmentService.NoAnnual);
            annual = limits.AnnualFull ? "FULL" : Money(limits.AnnualLimit);
        }
        if (limits is not null) visit = limits.VisitFull ? "FULL" : Money(limits.VisitLimit);
        var spec = SpecLimit(await repository.SpecLimitAsync(s, plan, pay, ins), warnings, mbmType == "S");
        return (annual, visit, IsNum(spec) ? Money(spec) : spec);
    }

    /// <summary>showMBMCov + lstCovAdd_Click: each covered person with the limits the desktop lists for it.</summary>
    private async Task<List<EnqCover>> Covers(IDbSession s, List<Dictionary<string, object?>> rows, Dictionary<string, object?> p, string number, EnqFlags flags,
        bool shared, string limitStatus, string annual, string visit, string spec, List<EnqClaimRow> claims, EnqPlan pl, List<string> warnings)
    {
        string ins = T(p, "INSType"), plan = T(p, "PLNType"), pay = T(p, "PAYCode");
        DateTime? eff = p.GetValueOrDefault("MBMPolEffDate") as DateTime?, exp = p.GetValueOrDefault("MBMPolExpDate") as DateTime?;
        var refunds = T(p, "MBMCLNStatus") == "C" ? await adjustment.RefundsAsync(s, number) : [];
        var sDentalLimit = pl.AnnualFound ? pl.SDentalAnnLimit ?? 0 : 0;
        var clickStatus = eff is null ? "" : (await repository.CoverStatusAsync(s, plan, ins, eff.Value)).Annual is { Count: > 0 } ca ? Blank(ca[^1], "A") : "";
        var shareTaken = false;
        var dentalRows = new Dictionary<string, List<EnqDentalRow>>();
        async Task<List<EnqDentalRow>> Dental(string cov) => dentalRows.TryGetValue(cov, out var d) ? d : dentalRows[cov] = await repository.DentalAsync(s, number, cov);
        async Task<(decimal Topup, decimal Sp)> Topup(string cov) => await repository.TopupAsync(s, number, cov) ?? (0m, 0m);

        var list = new List<EnqCover>();
        foreach (var c in rows.Where(r => T(r, "MBMCCoverID") is { Length: > 0 } id && id != "00"))
        {
            string cov = T(c, "MBMCCoverID"), coverPlan = T(c, "MBMCCLNPLNCode") is { Length: > 0 } cp ? cp : plan, rel = T(c, "MBMCRELCode");
            (List<string?> annualRows, List<string?> specRows) = eff is null ? (new List<string?>(), new List<string?>()) : await repository.CoverStatusAsync(s, coverPlan, ins, eff.Value);
            string status = annualRows.Count > 0 ? Blank(annualRows[^1], "") : "", specStatus = specRows.Count > 0 ? Blank(specRows[^1], "A") : "";

            if (T(p, "ASOInd") == "Y" || status != "F")
                (annual, visit, spec) = await CallLimit(s, p, number, cov, "S", ins, plan, pay, eff, limitStatus, warnings);
            else if (await repository.FLimitAsync(s, plan) is { } fl)
            {
                if (rel is "W" or "H") annual = fl.AnnLmtIndSupp == "Y" ? Cstr(fl.SuppLimit, "") : "FULL";
                else if (rel is "S" or "D") annual = fl.AnnLmtIndSupp == "Y" ? Cstr(fl.ChildDepLimit ?? 0, "0") : "FULL";
                (spec, visit) = (fl.SSpecLimit ?? "", fl.VisLmtInd ?? "");
            }

            decimal ffs, acc;
            var specBenefit = new EnqBenefit(NA, NA, NA);
            var spTopup = (await Topup(cov)).Sp;
            if (spec.Trim() == NA) (ffs, acc) = CoverGp(flags, status, CoverFilter(claims, CoverMode(status, cov, false), cov));
            else
            {
                (ffs, acc) = CoverGpWithSpec(status, CoverFilter(claims, CoverMode(status, cov, true), cov));
                var (_, specAcc) = CoverSpec(specStatus, status, cov, CoverFilter(claims, CoverMode(specStatus, cov, true), cov));
                var counts = specStatus != "A" && ((specStatus == "B" && cov == "01") || specStatus is "C" or "" || (specStatus == "D" && cov == "01"));
                specBenefit = IsNum(spec)
                    ? counts ? new EnqBenefit(Money(spec), Money(specAcc), Money(Num(spec) + spTopup - specAcc)) : Zero
                    : spec.Trim() == "FULL" ? new EnqBenefit("FULL", Money(specAcc), "FULL") : Zero;
            }

            var gpCounts = status != "A" && !shared;
            var accLabel = gpCounts && (status is "B" or "C" or "F" or "" or "E" || (status == "D" && cov == "01")) ? Money(acc) : "0.00";
            string limitLabel = "0.00", balLabel = "0.00";
            if (IsNum(annual))
            {
                var topup = (await Topup("01")).Topup;
                var take = gpCounts && ((status == "B" && !shareTaken) || status is "C" or "F" or "" or "E" || (status == "D" && cov == "01")
                                        || (status == "D" && cov == "02" && !shareTaken));
                if (take)
                {
                    (limitLabel, balLabel) = (Money(annual), Money(Num(annual) - acc + topup));
                    if (status is "B" or "D") shareTaken = true;
                }
            }
            else if (annual.Trim() == "FULL") (limitLabel, balLabel) = ("FULL", "FULL");
            var gp = shared ? new EnqBenefit(NA, NA, NA) : new EnqBenefit(limitLabel, accLabel, balLabel);

            var opticalCover = SuppOpticalCover(flags, cov);
            var suppOptical = opticalCover is null ? 0 : SuppOptical(flags, await Dental(opticalCover));
            var dental = SuppDentalLabels(flags, cov, sDentalLimit, SuppDentalRead(flags) ? SuppDental(flags, await Dental(cov)) : (0m, 0m), suppOptical);

            // lstCovAdd_Click: FFS (CIS_FFS2013) and top-up by the principal plan's limit status.
            int? clickMode = clickStatus switch
            {
                "B" when cov == "01" => 2,
                "C" or "" => 3,
                "D" when cov == "01" => 2,
                "E" => 3,
                _ => null,
            };
            List<decimal?> ffsRows = clickStatus == "A" || clickMode is null ? [] : await repository.CoverFfsAsync(s, number, clickMode.Value, cov);
            var ffsLabel = ffsRows.Count == 0 ? "0.00" : ffsRows.Sum(x => x ?? 0).ToString("#,###.00", CultureInfo.InvariantCulture);
            var clickTopup = clickStatus switch
            {
                "B" or "D" when cov == "01" => Money((await Topup("01")).Topup),
                "C" or "" => Money((await Topup(cov)).Topup),
                _ => "0.00",
            };

            var dob = c.GetValueOrDefault("MBMCDOB") as DateTime?;
            var ageBand = eff is null || dob is null ? "" : await AgeBand(s, eff.Value, dob.Value);
            var refund = refunds.TryGetValue(cov, out var rf) ? rf : ((decimal, decimal, decimal, decimal)?)null;
            var coverEff = c.GetValueOrDefault("MBMCEffDate") as DateTime?;
            list.Add(new EnqCover(cov, T(c, "MBMCName"), T(c, "MBMCICBCPPNo"), rel, dob, coverPlan, T(c, "MBMCCLNStatus"), T(c, "MBMCCardType"), N(c, "MBMCCLNPrem"),
                N(c, "MBMCCLNMCO"), T(c, "MBMCCLNWarranty"), T(c, "MBMCCLNRemarks"), refund?.Item1, refund?.Item2, refund?.Item3, refund?.Item4, gp, specBenefit, dental,
                new EnqBenefit(NA, opticalCover is null ? NA : Money(suppOptical), NA), IsNum(visit) ? Money(visit) : visit, N(c, "MBMCCLNMEPLmt"), N(c, "MBMCCLNDentalLmt"),
                N(c, "MBMCCLNDentalExtLmt"), N(c, "MBMCCLNSpecLmt"), N(c, "MBMCCLNFirstSpecLmt"), ageBand, ageBand == "01" ? "C" : "A",
                coverEff is null ? eff : coverEff, coverEff is null ? exp : c.GetValueOrDefault("MBMCExpDate") as DateTime?, ffsLabel, clickTopup, Money(spTopup)));
        }
        return list;
    }

    private static readonly EnqBenefit Zero = new("0.00", "0.00", "0.00");

    private async Task<string> AgeBand(IDbSession s, DateTime eff, DateTime dob)
    {
        var age = MemberRegistrationService.GetAge(eff, dob);
        var code = await adjustment.AgeCodeAsync(s, int.Parse(age[..^1], CultureInfo.InvariantCulture), age.EndsWith('D'));
        return code.Length >= 2 ? code[..2] : code.PadLeft(2, '0');
    }

    private static IEnumerable<EnqClaimRow> CoverFilter(List<EnqClaimRow> rows, int? mode, string cov) => mode switch
    {
        2 => rows.Where(r => r.CoverId?.Trim() != "00"),
        3 => rows.Where(r => r.CoverId?.Trim() == cov),
        _ => [],
    };

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
        long l => l,
        bool b => b ? 1 : 0,
        short h => h,
        byte b => b,
        _ => 0,
    };
}
