using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.LgTemplates;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="ILgBindingRepository"/> on the <c>cis.LgBinding_*</c> procedures.</summary>
public sealed class LgBindingRepository : ILgBindingRepository
{
    /// <inheritdoc />
    public async Task<LgBindingLookups> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.LgBinding_Lookups");
        return new LgBindingLookups((await grid.ReadAsync<LookupOption>()).ToList(), (await grid.ReadAsync<LookupOption>()).ToList(),
            (await grid.ReadAsync<LookupOption>()).ToList(), (await grid.ReadAsync<LookupOption>()).ToList());
    }

    /// <inheritdoc />
    public async Task<LgGroupPlans> PlansOfGroupAsync(IDbSession session, string groupCode)
    {
        using var grid = await session.MultipleAsync("cis.LgBinding_PlansOfGroup", new { groupCode });
        return new LgGroupPlans((await grid.ReadAsync<LookupOption>()).ToList(), (await grid.ReadAsync<LookupOption>()).ToList());
    }

    /// <inheritdoc />
    public Task<List<LgBinding>> SearchAsync(IDbSession session, string groupName, string? payCode, string? planCode) =>
        session.QueryAsync<LgBinding>("cis.LgBinding_Search", new { groupName, payCode, planCode });

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, LgBindingSaveRequest r, string depInd) =>
        session.ExecAsync("cis.LgBinding_Insert", Parameters(r, depInd));

    /// <inheritdoc />
    public async Task<int> UpdatePlanAsync(IDbSession session, LgBindingSaveRequest r, string depInd) =>
        await session.ScalarAsync<int>("cis.LgBinding_UpdatePlan", Parameters(r, depInd));

    private static object Parameters(LgBindingSaveRequest r, string depInd) => new
    {
        groupName = r.GroupName, planCode = r.PlanCode, glType = r.GlType, template = r.Template, payCode = r.PayCode, lgFormat = r.LgFormat,
        vip = r.Vip, depInd, depTemplate = r.DependentTemplate, remarks = r.Remarks,
    };
}
