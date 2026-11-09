using Cis.Application.Abstractions;
using Cis.Application.Bordx;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Bordx;

/// <summary><see cref="IBordxBsnRepository"/> on cis.BordxBsn_Summary.</summary>
public sealed class BordxBsnRepository : IBordxBsnRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SummaryAsync(IDbSession session, string bordxNos) =>
        session.RowsAsync("cis.BordxBsn_Summary", new { bordxNos }, 300);
}
