using Cis.Application.Abstractions;
using Cis.Application.AcExport;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.AcExport;

/// <summary><see cref="IPostInvMbmRepository"/> on cis.SunPost_InvMbm_*.</summary>
public sealed class PostInvMbmRepository : IPostInvMbmRepository
{
    /// <inheritdoc />
    public Task<List<string>> CompaniesAsync(IDbSession session) => session.QueryAsync<string>("cis.SunPost_Companies");

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostInvMbmFilter f) =>
        session.RowsAsync("cis.SunPost_InvMbm_Search", new
        {
            batchNo = f.BatchNo, company = f.Company, invDateFrom = f.InvDateFrom, invDateTo = f.InvDateTo, caseFrom = f.CaseFrom, caseTo = f.CaseTo,
            posted = f.BatchNo is null ? f.Posted : null, postDateFrom = f.PostDateFrom, postDateTo = f.PostDateTo, batchFrom = f.BatchFrom, batchTo = f.BatchTo,
            bordxDateFrom = f.BordxDateFrom, bordxDateTo = f.BordxDateTo, bordxFrom = f.BordxFrom, bordxTo = f.BordxTo, regMonth = f.RegMonth, regYear = f.RegYear,
        }, 600);

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, PostInvMbmTick t, string accMonth, string accYear, string batchNo, string user) =>
        session.ExecAsync("cis.SunPost_InvMbm_Insert", new { casNumber = t.CasNumber, invoiceNo = t.BatchNo, accMonth, accYear, batchNo, user });
}
