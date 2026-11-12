using Cis.Application.Abstractions;
using Cis.Application.AcExport;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.AcExport;

/// <summary><see cref="IPostRecGrpRepository"/> on cis.SunPost_RecGrp_*.</summary>
public sealed class PostRecGrpRepository : IPostRecGrpRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostRecGrpFilter f) =>
        session.RowsAsync("cis.SunPost_RecGrp_Search", new
        {
            batchNo = f.BatchNo, company = f.Company, recFrom = f.ReceiptFrom, recTo = f.ReceiptTo, recDateFrom = f.ReceiptDateFrom, recDateTo = f.ReceiptDateTo,
            postDateFrom = f.PostDateFrom, postDateTo = f.PostDateTo, posted = f.Posted, batchFrom = f.BatchFrom, batchTo = f.BatchTo, sj = f.Sj,
        }, 1800);

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, string receiptNo, string accMonth, string accYear, string batchNo, string user) =>
        session.ExecAsync("cis.SunPost_RecGrp_Insert", new { receiptNo, accMonth, accYear, batchNo, user });
}
