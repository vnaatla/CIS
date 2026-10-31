using Cis.Application.Abstractions;
using Cis.Application.Audit;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Audit;

/// <summary><see cref="IActionLogRepository"/> on <c>cis.ActionLog_Insert</c> (HISDB.dbo.tblActionLog).</summary>
public sealed class ActionLogRepository : IActionLogRepository
{
    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, string userCode, int actionId, string notes) =>
        session.ExecAsync("cis.ActionLog_Insert", new { userId = userCode, actionId, notes });
}
