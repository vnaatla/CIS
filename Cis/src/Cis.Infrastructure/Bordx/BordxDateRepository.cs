using Cis.Application.Abstractions;
using Cis.Application.Bordx;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Bordx;

/// <summary><see cref="IBordxDateRepository"/> on cis.BordxDate_*.</summary>
public sealed class BordxDateRepository : IBordxDateRepository
{
    /// <inheritdoc />
    public async Task<List<BordxDateRow>> FindAsync(IDbSession session, string bordxNo) =>
        (await session.QueryAsync<BordxDateRow>("cis.BordxDate_Find", new { bordxNo })).ToList();

    /// <inheritdoc />
    public Task UpdateAsync(IDbSession session, string bordxNo, DateTime date, string user) =>
        session.ExecAsync("cis.BordxDate_Update", new { bordxNo, bordxDate = date, user });
}
