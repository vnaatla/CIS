using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Domain.Errors;

namespace Cis.Application.Invoice;

/// <summary>Medix.bas MemberSuppInfo's output (the member, plan and limit details of the invoice screens and FrmINVMBMinfo).</summary>
public sealed record MemberSupp
{
    /// <summary>Member / cover.</summary>
    public string MbmNumber { get; init; } = "";
    /// <summary>CoverId.</summary>
    public string CoverId { get; init; } = "";
    /// <summary>Patient name / IC / date of birth (the covered person's for a cover).</summary>
    public string PatientName { get; init; } = "";
    /// <summary>PatientIc.</summary>
    public string PatientIc { get; init; } = "";
    /// <summary>Date of birth.</summary>
    public DateTime? PatientDob { get; init; }
    /// <summary>Relationship code of a covered person.</summary>
    public string RelationType { get; init; } = "";
    /// <summary>Payor, HLT code, plan, insurance type.</summary>
    public string Payor { get; init; } = "";
    /// <summary>HltCode.</summary>
    public string HltCode { get; init; } = "";
    /// <summary>PlanCode.</summary>
    public string PlanCode { get; init; } = "";
    /// <summary>InsType.</summary>
    public string InsType { get; init; } = "";
    /// <summary>Policy.</summary>
    public string PolicyNo { get; init; } = "";
    /// <summary>Policy dates.</summary>
    public DateTime? EffDate { get; init; }
    /// <summary>Policy expiry.</summary>
    public DateTime? ExpDate { get; init; }
    /// <summary>Data / policy status.</summary>
    public string DataStatus { get; init; } = "";
    /// <summary>PolicyStatus.</summary>
    public string PolicyStatus { get; init; } = "";
    /// <summary>Balance and visit limits ("FULL" or an amount).</summary>
    public string BalanceLimit { get; init; } = "";
    /// <summary>VisitLimit.</summary>
    public string VisitLimit { get; init; } = "";
    /// <summary>Visit / annual indicators (PLN).</summary>
    public string VisitIndicator { get; init; } = "";
    /// <summary>AnnualIndicator.</summary>
    public string AnnualIndicator { get; init; } = "";
    /// <summary>Accumulated amount.</summary>
    public decimal AccumAmount { get; init; }
    /// <summary>Card type.</summary>
    public string CardType { get; init; } = "";
    /// <summary>Co-payment (CoPaymentMBM).</summary>
    public string CoPay { get; init; } = "";
    /// <summary>Dental / specialist / MEP status and limit status (principal P*, cover S*).</summary>
    public string DentalStatus { get; init; } = "";
    /// <summary>SpecStatus.</summary>
    public string SpecStatus { get; init; } = "";
    /// <summary>MepStatus.</summary>
    public string MepStatus { get; init; } = "";
    /// <summary>DentalLimitStatus.</summary>
    public string DentalLimitStatus { get; init; } = "";
    /// <summary>SpecLimitStatus.</summary>
    public string SpecLimitStatus { get; init; } = "";
    /// <summary>Limits.</summary>
    public string DentalLimit { get; init; } = "";
    /// <summary>SpecLimit.</summary>
    public string SpecLimit { get; init; } = "";
    /// <summary>MepLimit.</summary>
    public string MepLimit { get; init; } = "";
    /// <summary>DentalExtLimit.</summary>
    public string DentalExtLimit { get; init; } = "";
    /// <summary>FirstSpecLimit.</summary>
    public string FirstSpecLimit { get; init; } = "";
    /// <summary>Supplementary limit type (A default, B / C / D).</summary>
    public string SuppLimitType { get; init; } = "A";
    /// <summary>PLN benefit status, gross amount status, MM percentage, CLN contract, ASO indicator.</summary>
    public string BenefitStatus { get; init; } = "";
    /// <summary>GrossStatus.</summary>
    public string GrossStatus { get; init; } = "";
    /// <summary>MmPercent.</summary>
    public string MmPercent { get; init; } = "";
    /// <summary>ClinicContract.</summary>
    public string ClinicContract { get; init; } = "";
    /// <summary>AsoInd.</summary>
    public string AsoInd { get; init; } = "";
    /// <summary>MBMCLNLTM.</summary>
    public string ClnLtm { get; init; } = "";
    /// <summary>Group company flags.</summary>
    public string GroupCompany { get; init; } = "";
    /// <summary>GroupMcoStatus.</summary>
    public string GroupMcoStatus { get; init; } = "";
    /// <summary>McStatus.</summary>
    public string McStatus { get; init; } = "";
    /// <summary>PostageFees.</summary>
    public string PostageFees { get; init; } = "";
    /// <summary>BordxFees.</summary>
    public string BordxFees { get; init; } = "";
    /// <summary>Department and MC days balance.</summary>
    public string Department { get; init; } = "";
    /// <summary>MC days balance.</summary>
    public decimal McDaysBalance { get; init; }
    /// <summary>Cancel dates and the cover's clinic expiry.</summary>
    public string CancelDate { get; init; } = "";
    /// <summary>CoverCancelDate.</summary>
    public string CoverCancelDate { get; init; } = "";
    /// <summary>CoverClinicExpDate.</summary>
    public string CoverClinicExpDate { get; init; } = "";
}

/// <summary>Raw rows of <c>cis.MemberSupp_Info</c>.</summary>
/// <param name="Member">VIEW_MBM_Clinical.</param>
/// <param name="Group">GRPCompany.</param>
/// <param name="Clinical">MBMCLNOthers.</param>
/// <param name="Plan">PLN.</param>
/// <param name="CoPay">CoPaymentMBM.CoAmount.</param>
/// <param name="SuppLimitStatus">AnnualLimit.SuppLimitStatus.</param>
/// <param name="Cover">MBMCoveredPersons.</param>
/// <param name="CoverLimits">MBMCoveredPersonsCLN rows (the cover and cover 01).</param>
public record MemberSuppRows(Dictionary<string, object?>? Member, Dictionary<string, object?>? Group, Dictionary<string, object?>? Clinical,
    Dictionary<string, object?>? Plan, object? CoPay, string? SuppLimitStatus, Dictionary<string, object?>? Cover, List<Dictionary<string, object?>> CoverLimits);

/// <summary>Data access of MemberSuppInfo.</summary>
public interface IMemberSuppRepository
{
    /// <summary>The raw rows.</summary>
    Task<MemberSuppRows> RowsAsync(IDbSession session, string mbmNumber, string covId);
}

/// <summary>Member info of the invoice screens (MemberSuppInfo / FrmINVMBMinfo).</summary>
public interface IMemberSuppService
{
    /// <summary>The member's details.</summary>
    Task<MemberSupp> InfoAsync(string? mbmNumber, string? covId);
}

/// <inheritdoc />
public sealed class MemberSuppService(IDbSessions sessions, IMemberSuppRepository repository) : IMemberSuppService
{
    /// <summary>Desktop texts.</summary>
    public const string NotFound = "Membership No not found ! ", NoPlan = "Clinical Plan not found ! ", NoCover = "No such Cover Person ! ";

    /// <inheritdoc />
    public async Task<MemberSupp> InfoAsync(string? mbmNumber, string? covId)
    {
        var number = (mbmNumber ?? "").Trim().ToUpperInvariant();
        var cov = string.IsNullOrWhiteSpace(covId) ? "00" : covId.Trim();
        if (number.Length == 0) throw new NotFoundException(NotFound);
        await using var s = await sessions.OpenAsync();
        return Build(number, cov, await repository.RowsAsync(s, number, cov));
    }

    /// <summary>MemberSuppInfo's rules.</summary>
    public static MemberSupp Build(string number, string cov, MemberSuppRows r)
    {
        var m = r.Member ?? throw new NotFoundException(NotFound);
        var c = r.Clinical ?? new Dictionary<string, object?>();
        var p = r.Plan ?? throw new NotFoundException(NoPlan);
        var principal = cov == "00";
        var visSupp = T(p, "VISLmtIndSupp") is { Length: > 0 } v ? v : "Y";
        var annSupp = T(p, "ANNLmtIndSupp") is { Length: > 0 } a ? a : "Y";
        var suppType = annSupp == "Y" && r.SuppLimitStatus is { } st && st.Trim().Length > 0 ? st.Trim() : "A";
        var g = r.Group;
        var info = new MemberSupp
        {
            MbmNumber = number, CoverId = cov, PatientName = T(m, "MBMName"), PatientIc = T(m, "MBMIcBcPp"), PatientDob = m.GetValueOrDefault("MBMDOB") as DateTime?,
            Payor = T(c, "PAYCode"), HltCode = T(m, "HLTCode"), PlanCode = T(c, "PLNType"), InsType = T(c, "INSType"), PolicyNo = T(c, "MBMPolNumber"),
            EffDate = c.GetValueOrDefault("MBMPolEffDate") as DateTime?, ExpDate = c.GetValueOrDefault("MBMPolExpDate") as DateTime?,
            DataStatus = T(m, "MBMCLNDataStatus"), PolicyStatus = T(m, "MBMCLNStatus"),
            VisitIndicator = T(p, "VisLmtIND"),
            AnnualIndicator = principal && suppType == "D" ? "N" : T(p, "AnnLmtIND"),
            BalanceLimit = principal && suppType == "D" ? "FULL" : T(p, "AnnLmtIND") == "N" ? "FULL" : T(c, "MBMCLNBalLmt") is { Length: > 0 } b ? b : "0",
            VisitLimit = T(p, "VisLmtIND") == "N" ? "FULL" : T(c, "MBMCLNVisitLmt"),
            AccumAmount = N(c, "MBMAccumAmt"), CardType = T(c, "MBMCardType"), CoPay = r.CoPay is null ? "" : Convert.ToString(r.CoPay, CultureInfo.InvariantCulture)!,
            DentalStatus = T(p, principal ? "PDentalStatus" : "SDentalStatus"), SpecStatus = T(p, principal ? "PSpecStatus" : "SSpecStatus"),
            MepStatus = T(p, principal ? "PMEPStatus" : "SMEPStatus"), DentalLimitStatus = T(p, principal ? "PDentalLimit" : "SDentalLimit"),
            SpecLimitStatus = T(p, principal ? "PSpecLimit" : "SSpecLimit"),
            DentalLimit = T(c, "MBMDentalLmt"), SpecLimit = T(c, "MBMSpeclmt"), MepLimit = T(c, "MBMMEPLmt"), DentalExtLimit = T(c, "MBMDentalExtLmt"),
            FirstSpecLimit = T(c, "MBMFirstSpecLmt"), SuppLimitType = suppType, BenefitStatus = T(p, "BNFStatus"), GrossStatus = T(p, "GrossAmtStatus"),
            MmPercent = T(p, "PLNMMIndPerc") is { Length: > 0 } mm ? mm : "0", ClinicContract = T(p, "ClnContract"), AsoInd = T(p, "ASOInd"),
            ClnLtm = T(c, "MBMCLNLTM"), GroupCompany = T(m, "MBMGroupCompany"),
            GroupMcoStatus = g is null ? "" : T(g, "GRPMCOStatus"), McStatus = g is null ? "" : T(g, "GRPMCStatus"),
            PostageFees = g is null ? "" : T(g, "GRPPostageFees"), BordxFees = g is null ? "" : T(g, "GRPBordxFees"),
            Department = T(m, "MBMDepartment"), McDaysBalance = N(m, "MBMCLNMCDaysBal"), CancelDate = T(c, "MBMCancelDate"),
        };
        if (principal) return info;

        var person = r.Cover ?? throw new NotFoundException(NoCover);
        info = info with
        {
            PatientName = T(person, "MBMCName"), PatientDob = person.GetValueOrDefault("MBMCDOB") as DateTime?, PatientIc = T(person, "MBMCICBCPPNo"),
            RelationType = T(person, "MBMCRELCode"),
        };
        var limitCover = suppType is "B" or "D" ? "01" : cov;
        var l = r.CoverLimits.FirstOrDefault(x => T(x, "MBMCCoverID") == limitCover);
        if (l is null) return info with { DataStatus = "D", PolicyStatus = "D" };
        if (suppType is "B" or "C" or "D")
            info = info with
            {
                BalanceLimit = annSupp == "N" ? "FULL" : Num(l, "MBMCCLNBalLmt"), VisitLimit = visSupp == "N" ? "FULL" : Num(l, "MBMCCLNVisitLmt"),
                DentalLimit = T(l, "MBMCCLNDentalLmt"), SpecLimit = T(l, "MBMCCLNSpecLmt"), MepLimit = T(l, "MBMCCLNMEPLmt"),
                DentalExtLimit = T(l, "MBMCCLNDentalExtLmt"), FirstSpecLimit = T(l, "MBMCCLNFirstSpecLmt"), AccumAmount = N(l, "MBMCAccumAmt"),
            };
        return info with
        {
            CardType = T(l, "MBMCCardType"), DataStatus = T(l, "MBMCCLNStatus"), PolicyStatus = T(l, "MBMCCLNStatus"), CoverCancelDate = T(l, "MBMCCancelDate"),
            CoverClinicExpDate = T(l, "MBMCCLNExpDate"),
        };
    }

    private static string T(Dictionary<string, object?> r, string k) => r.GetValueOrDefault(k) switch
    {
        null => "",
        DateTime d => d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
        var v => Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim(),
    };

    private static decimal N(Dictionary<string, object?> r, string k) =>
        decimal.TryParse(T(r, k), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;

    private static string Num(Dictionary<string, object?> r, string k) => N(r, k).ToString(CultureInfo.InvariantCulture);
}
