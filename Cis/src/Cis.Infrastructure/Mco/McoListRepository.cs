using System.Data;
using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Mco;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Mco;

/// <summary><see cref="IMcoListRepository"/> on the cis.McoList_* procedures.</summary>
public sealed class McoListRepository : IMcoListRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, McoListFilter f) =>
        session.RowsAsync("cis.McoList_Search", new
        {
            mode = f.Mode, bordxDate = f.BordxDate, batchNo = f.BatchNo, from = f.From, to = f.To, invoiceNo = f.InvoiceNo, invoiced = f.Invoiced,
            payor = f.Payor, groupCode = f.GroupCode,
        }, 300);

    /// <inheritdoc />
    public async Task<(List<(string Code, string Name)> Payors, List<(string Code, string Name)> Groups)> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.McoList_Lookups");
        var pay = (await grid.ReadAsync<(string, string)>()).ToList();
        var grp = (await grid.ReadAsync<(string, string)>()).ToList();
        return (pay, grp);
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> CompanyAsync(IDbSession session, string code) =>
        (await session.RowsAsync("cis.McoList_Company", new { code })).FirstOrDefault();

    /// <inheritdoc />
    public async Task<string> NextInvoiceNoAsync(IDbSession session)
    {
        var p = new DynamicParameters();
        p.Add("invoiceNo", dbType: DbType.String, size: 10, direction: ParameterDirection.Output);
        await session.Sql().Connection.ExecuteAsync("cis.McoList_NextInvoiceNo", p, session.Sql().Transaction, commandType: CommandType.StoredProcedure);
        return p.Get<string>("invoiceNo");
    }

    /// <inheritdoc />
    public Task StampAsync(IDbSession session, string member, string cover, string invoiceNo, string user, decimal mco) =>
        session.ExecAsync("cis.McoList_StampInvoice", new { member, cover, invoiceNo, user, mco });
}
