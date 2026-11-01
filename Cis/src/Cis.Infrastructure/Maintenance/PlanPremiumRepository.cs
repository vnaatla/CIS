using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.Plan;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IPlanPremiumRepository"/> on the <c>cis.PlanPremium_*</c> procedures.</summary>
public sealed class PlanPremiumRepository : IPlanPremiumRepository
{
    /// <inheritdoc />
    public async Task<PremiumSetup> SetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.PlanPremium_Setup");
        async Task<List<LookupOption>> Next() => (await grid.ReadAsync<LookupOption>()).ToList();
        return new PremiumSetup(await Next(), await Next(), await Next(), await Next(), await Next());
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? ins, string? pay, string? hlt, string? age, string? pln) =>
        session.RowsAsync("cis.PlanPremium_Search", new { insCode = ins, payCode = pay, hltCode = hlt, ageCode = age, plnCode = pln });

    /// <inheritdoc />
    public Task<List<PremiumBand>> BandsAsync(IDbSession session, string ins, string pln) =>
        session.QueryAsync<PremiumBand>("cis.PlanPremium_Bands", new { insCode = ins, plnCode = pln });

    /// <inheritdoc />
    public Task<bool> ExistsAsync(IDbSession session, PremiumRow k, string? payCode) =>
        session.ScalarAsync<bool>("cis.PlanPremium_Exists", new { insCode = k.InsCode, hltCode = k.HltCode, plnCode = k.PlnCode, ageCode = k.AgeCode, effDate = k.EffDate, payCode });

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, PremiumRow r, string userCode, DateTime today) =>
        session.ExecAsync("cis.PlanPremium_Insert", new
        {
            insCode = r.InsCode, payCode = r.PayCode, hltCode = r.HltCode, ageCode = r.AgeCode, plnCode = r.PlnCode, amount = r.Amount,
            suppAmount = r.SuppAmount, suppStatus = r.SuppStatus, effDate = r.EffDate, version = r.Version, user = userCode, today,
        });

    /// <inheritdoc />
    public Task<bool> InUseAsync(IDbSession session, string ins, string hlt, string pln, string age, string? pay) =>
        session.ScalarAsync<bool>("cis.PlanPremium_InUse", new { insCode = ins, hltCode = hlt, plnCode = pln, ageCode = age, payCode = pay });

    /// <inheritdoc />
    public Task<int> DeleteAsync(IDbSession session, int prmCode) => session.ScalarAsync<int>("cis.PlanPremium_Delete", new { prmCode });
}
