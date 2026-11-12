using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="ITrkGrpClnRepository"/> on cis.TrkGrpCln_*.</summary>
public sealed class TrkGrpClnRepository : ITrkGrpClnRepository
{
    /// <inheritdoc />
    public async Task<(List<string> Companies, List<string> Clinics)> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.TrkGrpCln_Lookups");
        var companies = (await grid.ReadAsync<string>()).ToList();
        return (companies, (await grid.ReadAsync<string>()).ToList());
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, TrkGrpClnRequest r) =>
        session.RowsAsync("cis.TrkGrpCln_Search", new { mode = r.Mode, grpCode = r.GroupCode, clnCode = r.ClinicCode, company = r.GroupCompany, clinic = r.Clinic }, 300);
}
