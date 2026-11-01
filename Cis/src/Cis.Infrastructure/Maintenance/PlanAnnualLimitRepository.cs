using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.Plan;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IPlanAnnualLimitRepository"/> on the <c>cis.PlanAnnual_*</c> procedures.</summary>
public sealed class PlanAnnualLimitRepository : IPlanAnnualLimitRepository
{
    /// <inheritdoc />
    public async Task<AnnualLimitSetup> SetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.PlanAnnual_Setup");
        async Task<List<LookupOption>> Next() => (await grid.ReadAsync<LookupOption>()).ToList();
        return new AnnualLimitSetup(await Next(), await Next(), await Next(), await Next(), await Next());
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? pln, string? hlt, string? ins, string? pay, string? grp) =>
        session.RowsAsync("cis.PlanAnnual_Search", new { plnCode = pln, hltCode = hlt, insCode = ins, payCode = pay, grpCompany = grp });

    /// <inheritdoc />
    public Task<bool> ExistsAsync(IDbSession session, AnnualLimitRow r) =>
        session.ScalarAsync<bool>("cis.PlanAnnual_Exists", new { insCode = r.InsCode, hltCode = r.HltCode, plnCode = r.PlnCode, version = r.Version, effDate = r.EffDate });

    /// <inheritdoc />
    public Task<int> SaveAsync(IDbSession session, AnnualLimitRow r, string userCode, DateTime today) =>
        session.ScalarAsync<int>("cis.PlanAnnual_Save", new
        {
            index = r.Index, plnCode = r.PlnCode, insCode = r.InsCode, hltCode = r.HltCode, annLimit = r.AnnLimit, pDental = r.PDental, pSpec = r.PSpec,
            pMep = r.PMep, status = r.Status, suppLimit = r.SuppLimit, sDental = r.SDental, sSpec = r.SSpec, sMep = r.SMep, effDate = r.EffDate,
            version = r.Version, user = userCode, today,
        });

    /// <inheritdoc />
    public Task<bool> InUseAsync(IDbSession session, string plnCode) => session.ScalarAsync<bool>("cis.PlanAnnual_InUse", new { plnCode });

    /// <inheritdoc />
    public Task<int> DeleteAsync(IDbSession session, int index) => session.ScalarAsync<int>("cis.PlanAnnual_Delete", new { index });
}
