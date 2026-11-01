using Cis.Application.Abstractions;
using Cis.Domain.Errors;

namespace Cis.Application.Membership;

/// <summary>frmSearchMBM's option buttons (Mem Num, Policy Num, Mem Name, IC Num, Employee Num, Policy Sub Num) and frmSearchMBMIC.</summary>
public enum MemberSearchBy
{
    /// <summary>Mem Num: MBMNumber =.</summary>
    MemberNo,
    /// <summary>Policy Num: MBMPolicyNo =.</summary>
    PolicyNo,
    /// <summary>Mem Name: MBMName LIKE '%text%' (the default option).</summary>
    Name,
    /// <summary>IC Num: MBMIcBcPp =; 12+ characters also match the IC with / without dashes.</summary>
    IC,
    /// <summary>frmSearchMBMIC: MBMIcBcPp = the patient IC as typed.</summary>
    IcExact,
    /// <summary>Employee Num: MBMEmployeeNo =.</summary>
    EmployeeNo,
    /// <summary>Policy Sub Num: MBMPOLSubNo1 =.</summary>
    PolicySubNo,
}

/// <summary>One list row of the member search pop-ups.</summary>
/// <param name="MbmNumber">Member No.</param>
/// <param name="CoverId">ID (cover id; 00 = principal).</param>
/// <param name="Name">Member Name.</param>
/// <param name="Ic">IC No.</param>
/// <param name="PolicyNo">Policy No (frmSearchMBM05).</param>
/// <param name="PolicyStatus">PolSt / Status.</param>
/// <param name="DataStatus">DataSt.</param>
/// <param name="EffDate">Eff Date.</param>
/// <param name="ExpDate">Exp Date.</param>
/// <param name="GroupCompany">Group Company.</param>
/// <param name="Department">Dept.</param>
/// <param name="Supp">True for a supplementary member (MBMCoveredPersons).</param>
public record MemberSearchRow(string MbmNumber, string? CoverId, string? Name, string? Ic, string? PolicyNo, string? PolicyStatus, string? DataStatus,
    DateTime? EffDate, DateTime? ExpDate, string? GroupCompany, string? Department, bool Supp);

/// <summary>The list plus the desktop's counters (Principal, Supp) and whether a cap cut the result.</summary>
/// <param name="Rows">Principals first, then supplementary members.</param>
/// <param name="Principals">Principal count.</param>
/// <param name="Supps">Supplementary count.</param>
/// <param name="Capped">True when one of the two searches reached <see cref="MemberSearchService.MaxRows"/>.</param>
public record MemberSearchResult(List<MemberSearchRow> Rows, int Principals, int Supps, bool Capped);

/// <summary>Data access of the member search pop-ups (<c>cis.MemberSearch_*</c>).</summary>
public interface IMemberSearchRepository
{
    /// <summary>Search_MBMCLN2013 with parameters.</summary>
    Task<List<MemberSearchRow>> PrincipalsAsync(IDbSession session, MemberSearchBy by, string text, int max);

    /// <summary>Search_MBMCovCLN2013 with parameters.</summary>
    Task<List<MemberSearchRow>> SuppsAsync(IDbSession session, MemberSearchBy by, string text, int max);

    /// <summary>Search_CISMBM_05 + Search_CISMBMSupp05 (member numbers starting with the text).</summary>
    Task<List<MemberSearchRow>> ByPrefixAsync(IDbSession session, string prefix);
}

/// <summary>The member search pop-ups.</summary>
public interface IMemberSearchService
{
    /// <summary>frmSearchMBM (Search) and frmSearchMBMIC (<see cref="MemberSearchBy.IcExact"/>, run on open).</summary>
    Task<MemberSearchResult> SearchAsync(MemberSearchBy by, string? text);

    /// <summary>frmSearchMBM05: run on open with the caller's member number.</summary>
    Task<MemberSearchResult> ByPrefixAsync(string? memberNo);
}

/// <inheritdoc />
public sealed class MemberSearchService(IDbSessions sessions, IMemberSearchRepository repository) : IMemberSearchService
{
    /// <summary>Most rows of each search (the desktop streamed every match; a name search runs over ~1 million members).</summary>
    public const int MaxRows = 5000;

    /// <summary>Desktop texts.</summary>
    public const string TooShort = "Search String must be at least 3 characters!", NotFound = "No record found!", NotFound05 = "No record found.";

    /// <inheritdoc />
    public async Task<MemberSearchResult> SearchAsync(MemberSearchBy by, string? text)
    {
        var value = (text ?? "").Trim();
        if (by != MemberSearchBy.IcExact && value.Length < 3) throw new BusinessRuleException(TooShort);
        if (value.Length == 0) throw new NotFoundException(NotFound);
        await using var s = await sessions.OpenAsync();
        // cmdSearch_Click: principals; supplementary members only for Mem Num, Mem Name and IC Num (frmSearchMBMIC: both).
        var principals = (await repository.PrincipalsAsync(s, by, value, MaxRows)).Select(DesktopPrincipalColumns).ToList();
        var supps = by is MemberSearchBy.MemberNo or MemberSearchBy.Name or MemberSearchBy.IC or MemberSearchBy.IcExact
            ? await repository.SuppsAsync(s, by, value, MaxRows)
            : [];
        if (principals.Count + supps.Count == 0) throw new NotFoundException(NotFound);
        return new MemberSearchResult([.. principals, .. supps], principals.Count, supps.Count, principals.Count >= MaxRows || supps.Count >= MaxRows);
    }

    /// <inheritdoc />
    public async Task<MemberSearchResult> ByPrefixAsync(string? memberNo)
    {
        // Search_CISMBM_05 takes nvarchar(20): the text plus "%".
        var prefix = (memberNo ?? "").Trim();
        if (prefix.Length > 19) prefix = prefix[..19];
        await using var s = await sessions.OpenAsync();
        var rows = await repository.ByPrefixAsync(s, prefix);
        if (rows.Count == 0) throw new NotFoundException(NotFound05);
        return new MemberSearchResult(rows, rows.Count(r => !r.Supp), rows.Count(r => r.Supp), false);
    }

    /// <summary>MBMInquiryList fills only Member No, ID, Name, IC and the two dates for principals (PolSt, DataSt, Group Company and Dept
    /// stay blank), as the desktop list does.</summary>
    public static MemberSearchRow DesktopPrincipalColumns(MemberSearchRow r) =>
        r with { PolicyStatus = null, DataStatus = null, GroupCompany = null, Department = null };
}
