using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="IClmTrkRepository"/> on cis.ClmTrk_Run.</summary>
public sealed class ClmTrkRepository : IClmTrkRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> RunAsync(IDbSession session, int caseNo, IReadOnlyDictionary<string, object?> filters)
    {
        var p = new DynamicParameters(filters);
        p.Add("case", caseNo);
        return session.RowsAsync("cis.ClmTrk_Run", p, 600);
    }

    /// <inheritdoc />
    public Task<List<string>> LookupAsync(IDbSession session, string kind, string? key) => session.QueryAsync<string>("cis.ClmTrk_Lookup", new { kind, key });
}
