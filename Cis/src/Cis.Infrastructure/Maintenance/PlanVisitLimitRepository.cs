using Cis.Application.Abstractions;
using Cis.Application.Maintenance.Plan;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IPlanVisitLimitRepository"/> on the <c>cis.PlanVisit_*</c> procedures.</summary>
public sealed class PlanVisitLimitRepository : IPlanVisitLimitRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? pln, string? hlt, string? ins, string? pay, string? grp) =>
        session.RowsAsync("cis.PlanVisit_Search", new { plnCode = pln, hltCode = hlt, insCode = ins, payCode = pay, grpCompany = grp });

    /// <inheritdoc />
    public Task<bool> ExistsAsync(IDbSession session, VisitLimitRow r) =>
        session.ScalarAsync<bool>("cis.PlanVisit_Exists", new { insCode = r.InsCode, hltCode = r.HltCode, plnCode = r.PlnCode, version = r.Version, effDate = r.EffDate });

    /// <inheritdoc />
    public Task<int> SaveAsync(IDbSession session, VisitLimitRow r, string userCode, DateTime today) =>
        session.ScalarAsync<int>("cis.PlanVisit_Save", new
        {
            index = r.Index, plnCode = r.PlnCode, insCode = r.InsCode, hltCode = r.HltCode, visitLimit = r.VisitLimit, status = r.Status,
            suppLimit = r.SuppLimit, effDate = r.EffDate, version = r.Version, user = userCode, today,
        });

    /// <inheritdoc />
    public Task<int> DeleteAsync(IDbSession session, int index) => session.ScalarAsync<int>("cis.PlanVisit_Delete", new { index });
}
