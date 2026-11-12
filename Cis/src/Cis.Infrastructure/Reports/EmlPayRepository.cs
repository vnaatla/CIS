using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="IEmlPayRepository"/> on cis.EmlPay_*.</summary>
public sealed class EmlPayRepository : IEmlPayRepository
{
    /// <inheritdoc />
    public Task<List<string>> CompaniesAsync(IDbSession session) => session.QueryAsync<string>("cis.EmlPay_Companies");

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, DateTime from, DateTime to, string? pvFrom, string? pvTo, bool resend, string? company,
        bool cs) =>
        session.RowsAsync("cis.EmlPay_Search", new { pvDateFrom = from, pvDateTo = to, pvFrom, pvTo, resend, company, cs }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> DetailsAsync(IDbSession session, string pvNo) => session.RowsAsync("cis.EmlPay_Details", new { pvNo });

    /// <inheritdoc />
    public Task MarkAsync(IDbSession session, string pvNo) => session.ExecAsync("cis.EmlPay_Mark", new { pvNo });
}
