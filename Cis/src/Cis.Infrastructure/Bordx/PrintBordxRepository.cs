using Cis.Application.Abstractions;
using Cis.Application.Bordx;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Bordx;

/// <summary><see cref="IPrintBordxRepository"/> on cis.PrintBordx_*.</summary>
public sealed class PrintBordxRepository : IPrintBordxRepository
{
    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> InfoAsync(IDbSession session, string bordxNo) =>
        (await session.RowsAsync("cis.PrintBordx_Info", new { bordxNo })).FirstOrDefault();

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, string bordxNo) => session.RowsAsync("cis.PrintBordx_Rows", new { bordxNo }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SummaryAsync(IDbSession session, string bordxNo, string variant) =>
        session.RowsAsync("cis.PrintBordx_Summary", new { bordxNo, variant }, 300);

    /// <inheritdoc />
    public Task UpdateAsync(IDbSession session, string bordxNo, string month, string year, string user) =>
        session.ExecAsync("cis.PrintBordx_Update", new { bordxNo, month, year, user });
}
