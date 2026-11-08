using System.Data;
using Cis.Application.Abstractions;
using Cis.Application.Mco;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Mco;

/// <summary><see cref="IMcoRefundRepository"/> on the cis.McoRefund_* procedures.</summary>
public sealed class McoRefundRepository : IMcoRefundRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> PrincipalsAsync(IDbSession session, McoRefundFilter f, string noted) =>
        session.RowsAsync("cis.McoRefund_Search", new { payor = f.Payor, company = f.Company, department = f.Department, byBordx = f.ByBordx, from = f.From, to = f.To, noted }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SupplementariesAsync(IDbSession session, McoRefundFilter f, IReadOnlyList<string> members) =>
        session.RowsAsync("cis.McoRefund_Supp", new
        {
            payor = f.Payor, company = f.Company, department = f.Department, byBordx = f.ByBordx, from = f.From, to = f.To, members = string.Join('|', members),
        }, 300);

    /// <inheritdoc />
    public async Task<(List<(string Code, string Name)> Payors, List<string> Companies, List<string> Departments)> LookupsAsync(IDbSession session, string? company)
    {
        using var grid = await session.MultipleAsync("cis.McoRefund_Lookups", new { company });
        var payors = (await grid.ReadAsync<(string, string)>()).ToList();
        var companies = (await grid.ReadAsync<string>()).ToList();
        var departments = (await grid.ReadAsync<string>()).ToList();
        return (payors, companies, departments);
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> CompanyAsync(IDbSession session, string? payor, string? company) =>
        (await session.RowsAsync("cis.McoRefund_Company", new { payor, company })).FirstOrDefault();

    /// <inheritdoc />
    public async Task<string> NextCreditNoteAsync(IDbSession session)
    {
        var p = new DynamicParameters();
        p.Add("creditNote", dbType: DbType.String, size: 10, direction: ParameterDirection.Output);
        await session.Sql().Connection.ExecuteAsync("cis.McoRefund_NextCreditNote", p, session.Sql().Transaction, commandType: CommandType.StoredProcedure);
        return p.Get<string>("creditNote");
    }

    /// <inheritdoc />
    public Task StampAsync(IDbSession session, string member, string creditNote, string user) =>
        session.ExecAsync("cis.McoRefund_StampCreditNote", new { member, creditNote, user });
}
