using System.Data;
using Cis.Application.Abstractions;
using Cis.Application.AcExport;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.AcExport;

/// <summary><see cref="IInvPostRepository"/> on cis.InvPost_*.</summary>
public sealed class InvPostRepository : IInvPostRepository
{
    /// <inheritdoc />
    public async Task<(List<Dictionary<string, object?>> Summary, List<Dictionary<string, object?>> Detail)> SearchAsync(IDbSession session, string? payCode,
        string? invFrom, string? invTo, DateTime? regFrom, DateTime? regTo, string? posted)
    {
        using var grid = await session.Sql().Connection.QueryMultipleAsync("cis.InvPost_Search", new { payCode, invFrom, invTo, regFrom, regTo, posted },
            session.Sql().Transaction, 600, CommandType.StoredProcedure);
        var summary = RowMapper.ToDictionaries(await grid.ReadAsync());
        return (summary, RowMapper.ToDictionaries(await grid.ReadAsync()));
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, string? payCode, string? invFrom, string? invTo, DateTime? regFrom, DateTime? regTo,
        string posted) =>
        session.RowsAsync("cis.InvPost_Rows", new { payCode, invFrom, invTo, regFrom, regTo, posted }, 600);

    /// <inheritdoc />
    public Task MarkAsync(IDbSession session, string invoiceNo, string clinicCode, string user) =>
        session.ExecAsync("cis.InvPost_Mark", new { invoiceNo, clinicCode, user });
}
