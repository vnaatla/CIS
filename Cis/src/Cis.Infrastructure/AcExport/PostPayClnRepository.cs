using Cis.Application.Abstractions;
using Cis.Application.AcExport;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.AcExport;

/// <summary><see cref="IPostPayClnRepository"/> on cis.SunPost_Pay2Cln_*.</summary>
public sealed class PostPayClnRepository : IPostPayClnRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostPayClnFilter f) =>
        session.RowsAsync("cis.SunPost_Pay2Cln_Search", new
        {
            done = false, batchNo = f.BatchNo, clinic = f.Clinic, pvDateFrom = f.PvDateFrom, pvDateTo = f.PvDateTo, pvFrom = f.PvFrom, pvTo = f.PvTo,
            posted = f.BatchNo is null ? f.Posted : null, postDateFrom = f.PostDateFrom, postDateTo = f.PostDateTo, batchFrom = f.BatchFrom, batchTo = f.BatchTo, sj = f.Sj,
        }, 600);

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, PostPayClnTick t, string accMonth, string accYear, string batchNo, string user) =>
        session.ExecAsync("cis.SunPost_Pay2Cln_Insert", new { pvNo = t.PvNo, invoiceNo = t.Analysis, accMonth, accYear, batchNo, user });
}
