using System.Data;
using Cis.Application.Abstractions;
using Cis.Application.AcExport;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.AcExport;

/// <summary><see cref="IPostInvClnRepository"/> on cis.SunPost_*.</summary>
public sealed class PostInvClnRepository : IPostInvClnRepository
{
    /// <inheritdoc />
    public Task<List<string>> PeriodsAsync(IDbSession session) => session.QueryAsync<string>("cis.SunPost_Periods");

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostInvClnFilter f) =>
        session.RowsAsync("cis.SunPost_InvCln_Search", new
        {
            batchNo = f.BatchNo, clinic = f.Clinic, invDateFrom = f.InvDateFrom, invDateTo = f.InvDateTo, invFrom = f.InvoiceFrom, invTo = f.InvoiceTo,
            posted = f.BatchNo is null ? f.Posted : null, postDateFrom = f.PostDateFrom, postDateTo = f.PostDateTo, batchFrom = f.BatchFrom, batchTo = f.BatchTo,
            bordxMonthFrom = f.BordxMonthFrom, bordxMonthTo = f.BordxMonthTo, bordxYearFrom = f.BordxYearFrom, bordxYearTo = f.BordxYearTo, accMonth = f.AccMonth,
            accYear = f.AccYear, pvDateFrom = f.PvDateFrom, pvDateTo = f.PvDateTo,
        }, 600);

    /// <inheritdoc />
    public async Task<string> NextInvNoAsync(IDbSession session)
    {
        var p = new DynamicParameters();
        p.Add("seqNo", dbType: DbType.String, size: 20, direction: ParameterDirection.Output);
        await session.Sql().Connection.ExecuteAsync("cis.SunPost_NextInvNo", p, session.Sql().Transaction, commandType: CommandType.StoredProcedure);
        return p.Get<string>("seqNo");
    }

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, PostInvClnTick t, string accMonth, string accYear, string batchNo, string user) =>
        session.ExecAsync("cis.SunPost_InvCln_Insert", new
        {
            invoiceNo = t.InvoiceNo, clinic = t.ClinicCode, accMonth, accYear, batchNo, invDate = t.InvoiceDate, user,
        });
}
