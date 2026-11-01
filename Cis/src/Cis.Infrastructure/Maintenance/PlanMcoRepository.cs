using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.Plan;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IPlanMcoRepository"/> on the <c>cis.PlanMco_*</c> procedures.</summary>
public sealed class PlanMcoRepository : IPlanMcoRepository
{
    /// <inheritdoc />
    public async Task<McoSetup> SetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.PlanMco_Setup");
        return new McoSetup((await grid.ReadAsync<LookupOption>()).ToList(), (await grid.ReadAsync<LookupOption>()).ToList());
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? hlt, string? pln, DateTime? eff) =>
        session.RowsAsync("cis.PlanMco_Search", new { hltCode = hlt, plnCode = pln, effDate = eff });

    /// <inheritdoc />
    public Task<bool> ExistsAsync(IDbSession session, string pln, string hlt, DateTime eff, string ins) =>
        session.ScalarAsync<bool>("cis.PlanMco_Exists", new { plnCode = pln, hltCode = hlt, effDate = eff, insCode = ins });

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, string pln, string hlt, string supp, DateTime eff, McoRow r, string userCode, DateTime today) =>
        session.ExecAsync("cis.PlanMco_Insert", new
        {
            plnCode = pln, hltCode = hlt, insCode = r.InsCode, suppStatus = supp, effDate = eff, adult = r.Adult, child = r.Child,
            suppAdult = r.SuppAdult, suppChild = r.SuppChild, user = userCode, today,
        });

    /// <inheritdoc />
    public Task<int> DeleteAsync(IDbSession session, int index) => session.ScalarAsync<int>("cis.PlanMco_Delete", new { index });
}
