using Cis.Application.Abstractions;
using Cis.Application.Membership;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Membership;

/// <summary><see cref="IGeneralExclusionsRepository"/> on the <c>cis.GeneralExclusions_*</c> procedures.</summary>
public sealed class GeneralExclusionsRepository : IGeneralExclusionsRepository
{
    /// <inheritdoc />
    public Task<List<string>> GroupCompaniesAsync(IDbSession session) => session.QueryAsync<string>("cis.GeneralExclusions_GroupCompanies");

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> GetAsync(IDbSession session, string groupCompany) =>
        (await session.RowsAsync("cis.GeneralExclusions_Get", new { groupCompany })).FirstOrDefault();
}
