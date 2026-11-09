using Cis.Application.Abstractions;
using Cis.Application.Bordx;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Bordx;

/// <summary><see cref="IBordxRemoveRepository"/> on cis.BordxRemove_*.</summary>
public sealed class BordxRemoveRepository : IBordxRemoveRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ListAsync(IDbSession session, string bordxNo) => session.RowsAsync("cis.BordxRemove_List", new { bordxNo }, 300);

    /// <inheritdoc />
    public Task RemoveAsync(IDbSession session, string bordxNo, string casNumber, string user) =>
        session.ExecAsync("cis.BordxRemove_Case", new { bordxNo, casNumber, user });
}
