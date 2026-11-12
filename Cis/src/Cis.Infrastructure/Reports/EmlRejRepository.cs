using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="IEmlRejRepository"/> on cis.EmlRej_*.</summary>
public sealed class EmlRejRepository : IEmlRejRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, DateTime from, DateTime to, string? caseFrom, string? caseTo, bool resend,
        string? company) =>
        session.RowsAsync("cis.EmlRej_Search", new { regFrom = from, regTo = to, caseFrom, caseTo, resend, company }, 300);

    /// <inheritdoc />
    public Task MarkAsync(IDbSession session, string casNumber) => session.ExecAsync("cis.EmlRej_Mark", new { casNumber });
}
