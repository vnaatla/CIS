using Cis.Application.Abstractions;
using Cis.Application.AcExport;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.AcExport;

/// <summary><see cref="IPostPayMbmRepository"/> on cis.SunPost_Pay2Mbm_*.</summary>
public sealed class PostPayMbmRepository : IPostPayMbmRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostPayMbmFilter f) =>
        session.RowsAsync("cis.SunPost_Pay2Mbm_Search", new
        {
            batchNo = f.BatchNo, company = f.Company, pvDateFrom = f.PvDateFrom, pvDateTo = f.PvDateTo, pvFrom = f.PvFrom, pvTo = f.PvTo,
            posted = f.BatchNo is null ? f.Posted : null, postDateFrom = f.PostDateFrom, postDateTo = f.PostDateTo, batchFrom = f.BatchFrom, batchTo = f.BatchTo,
        }, 600);

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, string transRef, string accMonth, string accYear, string batchNo, string user) =>
        session.ExecAsync("cis.SunPost_Pay2Mbm_Insert", new { transRef, accMonth, accYear, batchNo, user });
}
