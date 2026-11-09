using System.Data;
using Cis.Application.Abstractions;
using Cis.Application.Bordx;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Bordx;

/// <summary><see cref="IBordxWsRepository"/> on the cis.BordxWs_* procedures.</summary>
public sealed class BordxWsRepository : IBordxWsRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, BordxWsFilter f)
    {
        object common = new
        {
            bordxType = f.BordxType, groupCode = f.GroupCode, scan = f.Scan, caseFrom = f.CaseFrom, caseTo = f.CaseTo, visitFrom = f.VisitFrom, visitTo = f.VisitTo,
            noVisitDate = f.NoVisitDate, department = f.Department, clinicType = f.ClinicType, opFrom = f.OpFrom, opTo = f.OpTo, subsidiary = f.Subsidiary,
            payTo = f.PayTo, treatment = f.Treatment, omni = f.Omni, regFrom = f.RegFrom, regTo = f.RegTo, clinic = f.Clinic, entity = f.Entity,
        };
        var p = new DynamicParameters(common);
        if (f.Uob is { } uob)
        {
            p.Add("flex", uob == 1);
            return session.RowsAsync("cis.BordxWs_SearchUob", p, 300);
        }
        p.Add("isPureAso", f.IsPureAso);
        return session.RowsAsync("cis.BordxWs_Search", p, 300);
    }

    /// <inheritdoc />
    public async Task<(List<(string Code, string Name)> Groups, List<(string Code, string Name)> Treatments, List<(string Code, string Name)> Clinics)> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.BordxWs_Lookups");
        return ((await grid.ReadAsync<(string, string)>()).ToList(), (await grid.ReadAsync<(string, string)>()).ToList(), (await grid.ReadAsync<(string, string)>()).ToList());
    }

    /// <inheritdoc />
    public async Task<(List<string> Departments, List<string> Subsidiaries)> CompanyAsync(IDbSession session, string company)
    {
        using var grid = await session.MultipleAsync("cis.BordxWs_Company", new { company });
        return ((await grid.ReadAsync<string>()).ToList(), (await grid.ReadAsync<string>()).ToList());
    }

    /// <inheritdoc />
    public async Task<bool> IsBtbgAsync(IDbSession session, string plnCode)
    {
        var p = new DynamicParameters(new { plnCode });
        p.Add("isBtbg", dbType: DbType.Boolean, direction: ParameterDirection.Output);
        await Exec(session, "cis.BordxWs_IsBtbg", p);
        return p.Get<bool?>("isBtbg") ?? false;
    }

    /// <inheritdoc />
    public async Task<string> NextNoAsync(IDbSession session, string payCode)
    {
        var p = new DynamicParameters(new { payCode });
        p.Add("bordxNo", dbType: DbType.String, size: 20, direction: ParameterDirection.Output);
        await Exec(session, "cis.BordxWs_NextNo", p);
        return p.Get<string>("bordxNo");
    }

    /// <inheritdoc />
    public async Task<decimal> UpdateCaseAsync(IDbSession session, string casNumber, string bordxNo, int mode, string user)
    {
        var p = new DynamicParameters(new { casNumber, bordxNo, mode, user });
        p.Add("amount", dbType: DbType.Currency, direction: ParameterDirection.Output);
        await Exec(session, "cis.BordxWs_UpdateCase", p);
        return p.Get<decimal?>("amount") ?? 0;
    }

    /// <inheritdoc />
    public Task SaveBordxAsync(IDbSession session, string bordxNo, string payCode, string user, decimal amount, int claims, string company, string rptType,
        string subsidiary) =>
        session.ExecAsync("cis.BordxWs_SaveBordx", new { bordxNo, payCode, user, amount, claims, company, rptType, subsidiary });

    private static Task Exec(IDbSession s, string proc, DynamicParameters p) =>
        s.Sql().Connection.ExecuteAsync(proc, p, s.Sql().Transaction, commandType: CommandType.StoredProcedure, commandTimeout: 120);
}
