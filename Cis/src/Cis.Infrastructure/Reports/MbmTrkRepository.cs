using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="IMbmTrkRepository"/> on cis.MbmTrk_Run.</summary>
public sealed class MbmTrkRepository : IMbmTrkRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> RunAsync(IDbSession session, int caseNo, IReadOnlyDictionary<string, object?> filters)
    {
        var p = new DynamicParameters(filters);
        p.Add("case", caseNo);
        return session.RowsAsync("cis.MbmTrk_Run", p, 600);
    }

    /// <inheritdoc />
    public Task<List<string>> LookupAsync(IDbSession session, string kind, string? key) => session.QueryAsync<string>("cis.MbmTrk_Lookup", new { kind, key });
}
