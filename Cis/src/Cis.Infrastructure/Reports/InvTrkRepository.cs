using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="IInvTrkRepository"/> on cis.InvTrk_*.</summary>
public sealed class InvTrkRepository : IInvTrkRepository
{
    /// <inheritdoc />
    public Task<List<string>> LookupAsync(IDbSession session, string kind, string? key) => session.QueryAsync<string>("cis.InvTrk_Lookup", new { kind, key });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, IReadOnlyDictionary<string, object?> parameters, bool jupiter) =>
        session.RowsAsync(jupiter ? "cis.InvTrkJup_Search" : "cis.InvTrk_Search", new DynamicParameters(parameters), 900);
}
