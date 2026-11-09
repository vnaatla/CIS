using Cis.Application.Abstractions;
using Cis.Application.Payment;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Payment;

/// <summary><see cref="IPayClnPostRepository"/> on cis.PayClnPost_*.</summary>
public sealed class PayClnPostRepository : IPayClnPostRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? payCode, string? invFrom, string? invTo, DateTime? regFrom, DateTime? regTo,
        string? posted) =>
        session.RowsAsync("cis.PayClnPost_Search", new { payCode, invFrom, invTo, regFrom, regTo, posted }, 300);

    /// <inheritdoc />
    public Task MarkAsync(IDbSession session, string invoiceNo, string clinicCode, string user) =>
        session.ExecAsync("cis.PayClnPost_Mark", new { invoiceNo, clinicCode, user });
}
