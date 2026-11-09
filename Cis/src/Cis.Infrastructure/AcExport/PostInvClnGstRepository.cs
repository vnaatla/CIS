using System.Data;
using Cis.Application.Abstractions;
using Cis.Application.AcExport;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.AcExport;

/// <summary><see cref="IPostInvClnGstRepository"/> on cis.SunPost_InvClnGST_* and cis.SunPost_NextBatch.</summary>
public sealed class PostInvClnGstRepository : IPostInvClnGstRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostInvClnGstFilter f) =>
        session.RowsAsync("cis.SunPost_InvClnGST_Search", new
        {
            batchNo = f.BatchNo, clinic = f.Clinic, invDateFrom = f.InvDateFrom, invDateTo = f.InvDateTo, invFrom = f.InvoiceFrom, invTo = f.InvoiceTo,
            posted = f.BatchNo is null ? f.Posted : null, postDateFrom = f.PostDateFrom, postDateTo = f.PostDateTo, batchFrom = f.BatchFrom, batchTo = f.BatchTo,
            sunFrom = f.SunFrom, sunTo = f.SunTo, accMonth = f.AccMonth, accYear = f.AccYear,
        }, 600);

    /// <inheritdoc />
    public async Task<string> NextBatchAsync(IDbSession session)
    {
        var p = new DynamicParameters();
        p.Add("batchNo", dbType: DbType.String, size: 20, direction: ParameterDirection.Output);
        await session.Sql().Connection.ExecuteAsync("cis.SunPost_NextBatch", p, session.Sql().Transaction, commandType: CommandType.StoredProcedure);
        return p.Get<string>("batchNo");
    }

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, PostInvClnTick t, string accMonth, string accYear, string invBatch, string batchNo, string user) =>
        session.ExecAsync("cis.SunPost_InvClnGST_Insert", new
        {
            invoiceNo = t.InvoiceNo, clinic = t.ClinicCode, accMonth, accYear, invBatch, invDate = t.InvoiceDate, batchNo, user,
        });
}
