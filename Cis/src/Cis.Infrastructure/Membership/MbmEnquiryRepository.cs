using Cis.Application.Abstractions;
using Cis.Application.Membership;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Membership;

/// <summary><see cref="IMbmEnquiryRepository"/> on <c>cis.MbmEnq_*</c>.</summary>
public sealed class MbmEnquiryRepository : IMbmEnquiryRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> MemberAsync(IDbSession session, string mbmNumber) =>
        session.RowsAsync("cis.MbmEnq_Member", new { mbmNumber });

    /// <inheritdoc />
    public async Task<EnqPlan> PlanAsync(IDbSession session, string mbmNumber, string plan, string insType, DateTime? expDate)
    {
        using var grid = await session.MultipleAsync("cis.MbmEnq_Plan", new { mbmNumber, plnCode = plan, insCode = insType, expDate });
        var annual = (await grid.ReadAsync<(string? SuppLimitStatus, decimal? PDentalAnnLimit, decimal? SDentalAnnLimit, decimal? PSpecAnnLimit, decimal? SSpecAnnLimit,
            decimal? AnnLimit)>()).ToList();
        var maternity = (await grid.ReadAsync<decimal?>()).ToList();
        var spec = (await grid.ReadAsync<string?>()).ToList();
        var pln = (await grid.ReadAsync<(string? MaternityInd, string? DentalInd)>()).ToList();
        var product = (await grid.ReadAsync<(string? ProductName, string? Description)>()).ToList();
        var flex = await grid.ReadSingleAsync<int>();
        return new EnqPlan(annual.Count > 0, annual.FirstOrDefault().SuppLimitStatus, annual.FirstOrDefault().PDentalAnnLimit, annual.FirstOrDefault().SDentalAnnLimit, annual.FirstOrDefault().PSpecAnnLimit,
            maternity.Count > 0, maternity.FirstOrDefault(), spec.Count > 0, spec.FirstOrDefault(), pln.FirstOrDefault().MaternityInd?.Trim() ?? "",
            pln.FirstOrDefault().DentalInd?.Trim() ?? "",
            product.Count > 0 ? $"{product[0].ProductName}({product[0].Description})" : "", flex);
    }

    /// <inheritdoc />
    public async Task<(decimal Topup, decimal SpTopup)?> TopupAsync(IDbSession session, string mbmNumber, string? covId)
    {
        var rows = (await session.QueryAsync<(decimal? TopupAmt, decimal? SpTopupAmt)>("cis.MbmEnq_Topup", new { mbmNumber, covId })).ToList();
        return rows.Count == 0 ? null : (rows[0].TopupAmt ?? 0, rows[0].SpTopupAmt ?? 0);
    }

    /// <inheritdoc />
    public async Task<(decimal Aso, decimal Op, string Combined)?> AsoAsync(IDbSession session, string mbmNumber, string covId)
    {
        var rows = await session.RowsAsync("cis.MbmEnq_Aso", new { mbmNumber, covId });
        if (rows.Count == 0) return null;
        var r = rows[^1];
        return (r["MBMASOLimit"] as decimal? ?? 0, r["MBMOPLimit"] as decimal? ?? 0, (r["IsIPOPCombined"] as string)?.Trim() is { Length: > 0 } c ? c : "N");
    }

    /// <inheritdoc />
    public async Task<EnqSpecLimit> SpecLimitAsync(IDbSession session, string plan, string payCode, string insType)
    {
        using var grid = await session.MultipleAsync("cis.MbmEnq_SpecLimit", new { plnCode = plan, payCode, insCode = insType });
        var ind = (await grid.ReadAsync<string?>()).ToList();
        var row = (await grid.ReadAsync<(decimal? PrinLimit, decimal? SuppLimit, string? SuppLimitStatus)>()).ToList();
        return new EnqSpecLimit(ind, row.Count > 0, row.FirstOrDefault().PrinLimit, row.FirstOrDefault().SuppLimitStatus?.Trim(), row.FirstOrDefault().SuppLimit);
    }

    /// <inheritdoc />
    public async Task<List<EnqClaimRow>> ClaimsAsync(IDbSession session, string mbmNumber) =>
        (await session.QueryAsync<EnqClaimRow>("cis.MbmEnq_Utilisation", new { mbmNumber, coverMode = 0 })).ToList();

    /// <inheritdoc />
    public async Task<List<EnqDentalRow>> DentalAsync(IDbSession session, string mbmNumber, string covId) =>
        (await session.QueryAsync<EnqDentalRow>("cis.MbmEnq_Dental", new { mbmNumber, covId })).ToList();

    /// <inheritdoc />
    public async Task<List<EnqDentalRow>> DentalOneLimitAsync(IDbSession session, string mbmNumber) =>
        (await session.QueryAsync<EnqDentalRow>("cis.MbmEnq_DentalOneLimit", new { mbmNumber })).ToList();

    /// <inheritdoc />
    public async Task<(List<string?> Annual, List<string?> Specialist)> CoverStatusAsync(IDbSession session, string plan, string insType, DateTime effDate)
    {
        using var grid = await session.MultipleAsync("cis.MbmEnq_CoverStatus", new { plnCode = plan, insCode = insType, effDate });
        return ((await grid.ReadAsync<string?>()).ToList(), (await grid.ReadAsync<string?>()).ToList());
    }

    /// <inheritdoc />
    public async Task<EnqFLimit?> FLimitAsync(IDbSession session, string plan) =>
        (await session.QueryAsync<EnqFLimit>("cis.MbmEnq_FLimit", new { plnCode = plan })).FirstOrDefault();

    /// <inheritdoc />
    public async Task<List<decimal?>> CoverFfsAsync(IDbSession session, string mbmNumber, int coverMode, string covId) =>
        await session.QueryAsync<decimal?>("cis.MbmEnq_CoverFfs", new { mbmNumber, coverMode, covId });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> FlexAsync(IDbSession session, string mbmNumber, int policyYear) =>
        session.RowsAsync("cis.MbmEnq_Flex", new { icno = (string?)null, empno = (string?)null, mbmno = mbmNumber, polYear = policyYear });

    /// <inheritdoc />
    public async Task<((decimal Limit, decimal Util, decimal Balance)? HomeOffice, decimal Principal, decimal Dependant)> UbNotesAsync(IDbSession session, string mbmNumber)
    {
        using var grid = await session.MultipleAsync("cis.MbmEnq_UbNotes", new { mbmNumber });
        var cat = (await grid.ReadAsync<(string? PLNType, int? PlncategoryInd, decimal? PlncategoryLimit, decimal? Util, decimal? Balance)>()).ToList();
        var util = (await grid.ReadAsync<(decimal? PriBordxamt, decimal? SuppBordxamt, string? MbmType, string? CoverID)>()).ToList();
        (decimal, decimal, decimal)? home = cat.Count == 0 ? null : (cat[^1].PlncategoryLimit ?? 0, cat[^1].Util ?? 0, cat[^1].Balance ?? 0);
        return (home, util.Where(u => u.CoverID?.Trim() == "00").Sum(u => u.PriBordxamt ?? 0), util.Where(u => u.CoverID?.Trim() != "00").Sum(u => u.SuppBordxamt ?? 0));
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> CasesAsync(IDbSession session, string mbmNumber, bool specialist) =>
        session.RowsAsync("cis.MbmEnq_Cases", new { mbmNumber, specialist });
}
