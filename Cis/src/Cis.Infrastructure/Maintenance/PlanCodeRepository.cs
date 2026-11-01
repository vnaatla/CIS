using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.Plan;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IPlanCodeRepository"/> on the <c>cis.PlanCode_*</c> procedures.</summary>
public sealed class PlanCodeRepository : IPlanCodeRepository
{
    /// <inheritdoc />
    public async Task<PlanCodeSetup> SetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.PlanCode_Setup");
        async Task<List<LookupOption>> Next() => (await grid.ReadAsync<LookupOption>()).ToList();
        return new PlanCodeSetup(await Next(), await Next(), await Next(), await Next(), await Next());
    }

    /// <inheritdoc />
    public Task<List<LookupOption>> PayorsAsync(IDbSession session, string hltCode) => session.QueryAsync<LookupOption>("cis.PlanCode_Payors", new { hltCode });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PlanCodeFilters f)
    {
        var p = new DynamicParameters(new
        {
            plnCode = f.PlnCode, payCode = f.PayCode, hisPayCode = f.HisPayCode, hltCode = f.HltCode, grpCode = f.GrpCode, clnContract = f.ClnContract,
        });
        foreach (var (name, value) in f.Flags) p.Add(name, value);
        return session.RowsAsync("cis.PlanCode_Search", p);
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> GetAsync(IDbSession session, string code) =>
        (await session.RowsAsync("cis.PlanCode_Get", new { code })).FirstOrDefault();

    /// <inheritdoc />
    public Task<bool> GroupNameExistsAsync(IDbSession session, string name) => session.ScalarAsync<bool>("cis.PlanCode_GroupNameExists", new { name });

    /// <inheritdoc />
    public async Task<string> NextCodeAsync(IDbSession session, string letter) => (await session.ScalarAsync<string>("cis.PlanCode_NextCode", new { letter }))!;

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, PlanRow r, string userCode, DateTime today) =>
        session.ExecAsync("cis.PlanCode_Insert", new
        {
            code = r.Code, desc = r.Desc, hltCode = r.HltCode, payCode = r.PayCode, hisPayCode = r.HisPayCode, grpCompany = r.GrpCompany,
            annLmt = r.AnnLmt, visLmt = r.VisLmt, sAnnLmt = r.SAnnLmt, sVisLmt = r.SVisLmt, prm = r.Prm, mco = r.Mco, clnContract = r.ClnContract,
            dcStatus = r.DcStatus, insured = r.Insured, hrApproval = r.HrApproval, opFamilyStatus = r.OpFamilyStatus, prinSpLimit = r.PrinSpLimit,
            prinGpLimit = r.PrinGpLimit, depSpLimit = r.DepSpLimit, depGpLimit = r.DepGpLimit, specialRemark = r.SpecialRemark,
            opDependent = r.OpDependent, user = userCode, today,
        });

    /// <inheritdoc />
    public Task<bool> InUseAsync(IDbSession session, string code) => session.ScalarAsync<bool>("cis.PlanCode_InUse", new { code });

    /// <inheritdoc />
    public Task<int> DeleteAsync(IDbSession session, string code) => session.ScalarAsync<int>("cis.PlanCode_Delete", new { code });
}
