using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="IEmlPrvRepository"/> on cis.EmlPrv_* and CISPayment sp_ProviderInv.</summary>
public sealed class EmlPrvRepository : IEmlPrvRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, DateTime from, DateTime to, string? pvFrom, string? pvTo, bool resend, bool cs) =>
        session.RowsAsync("cis.EmlPrv_Search", new { pvDateFrom = from, pvDateTo = to, pvFrom, pvTo, resend, cs }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> InvoicesAsync(IDbSession session, string pvNo) =>
        session.RowsAsync("CISPayment.dbo.sp_ProviderInv", new { PVCLNo = pvNo, CaseType = "C" }, 300);

    /// <inheritdoc />
    public Task MarkAsync(IDbSession session, string pvNo) => session.ExecAsync("cis.EmlPrv_Mark", new { pvNo });
}
