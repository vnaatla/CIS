using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="IClmTrk2Repository"/> on cis.ClmTrk2_Run.</summary>
public sealed class ClmTrk2Repository : IClmTrk2Repository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> RunAsync(IDbSession session, int caseNo, IReadOnlyDictionary<string, object?> filters)
    {
        var p = new DynamicParameters(filters);
        p.Add("case", caseNo);
        return session.RowsAsync("cis.ClmTrk2_Run", p, 600);
    }
}
