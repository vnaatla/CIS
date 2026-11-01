using Cis.Application.Abstractions;
using Cis.Application.Membership;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Membership;

/// <summary><see cref="IMbmTopupRepository"/>: HISDB SearchCISMemberProc and CISDB spUtilisationOP_test1 as the desktop calls them, plus <c>cis.MbmTopup_*</c>.</summary>
public sealed class MbmTopupRepository : IMbmTopupRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> MemberAsync(IDbSession session, string mbmNumber) =>
        session.RowsAsync("cis.Hdb_SearchCISMemberProc", new { MBMNumber = mbmNumber });

    /// <inheritdoc />
    public async Task<(string GroupCompany, string Department)> OthersAsync(IDbSession session, string mbmNumber)
    {
        var row = (await session.RowsAsync("cis.MbmTopup_Others", new { mbmNumber })).FirstOrDefault();
        return (row?["GroupCompany"] as string ?? "", row?["Department"] as string ?? "");
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> UtilisationAsync(IDbSession session, string mbmNumber) =>
        session.RowsAsync("dbo.spUtilisationOP_test1", new { CompanyCode = (string?)null, mbmnumber = mbmNumber, ICNo = (string?)null, EmpNo = (string?)null,
            depno = (string?)null, polno = (string?)null });

    /// <inheritdoc />
    public async Task<(string Code, string? Desc)?> LimitTypeAsync(IDbSession session, string plan, string insType, DateTime effDate)
    {
        var row = (await session.RowsAsync("cis.MbmTopup_LimitType", new { plnCode = plan, insCode = insType, effDate })).FirstOrDefault();
        return row is null ? null : ((string)row["LimitCode"]!, row["LimitDesc"] as string);
    }

    /// <inheritdoc />
    public async Task<string?> StepAsync(IDbSession session, string mbmNumber, bool next) =>
        (await session.RowsAsync("cis.MbmTopup_Step", new { mbmNumber, next })).FirstOrDefault()?["MBMNumber"] as string;

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, TopupSave t, string userCode, DateTime today) =>
        session.ExecAsync("cis.MbmTopup_Save", new
        {
            mbmNumber = t.MbmNumber, covId = t.CovId, gp = t.Gp, icNo = t.IcNo, prevIcNo = t.PrevIcNo, prevLimit = t.PrevLimit, topupAmt = t.TopupAmt,
            newLimit = t.NewLimit, limitCode = t.LimitCode, user = userCode, today,
        });
}
