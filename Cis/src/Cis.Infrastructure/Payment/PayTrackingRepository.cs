using Cis.Application.Abstractions;
using Cis.Application.Payment;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Payment;

/// <summary><see cref="IPayTrackingRepository"/> on cis.PayTracking_Search.</summary>
public sealed class PayTrackingRepository : IPayTrackingRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayTrackingFilter f)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        var r = f.Request;
        return session.RowsAsync("cis.PayTracking_Search", new
        {
            member = f.Member, by = f.By, caseFrom = B(r.CaseFrom), caseTo = B(r.CaseTo), company = B(r.Company), mbmNo = B(r.MbmNo), invFrom = B(r.InvoiceFrom),
            invTo = B(r.InvoiceTo), bordxFrom = B(r.BordxFrom), bordxTo = B(r.BordxTo), opFrom = B(r.OpFrom), opTo = B(r.OpTo), clinic = f.Clinic, invDateFrom = f.Dates[0],
            invDateTo = f.Dates[1], payment = r.Payment, pvFrom = B(r.PvFrom), pvTo = B(r.PvTo), pvDateFrom = f.Dates[2], pvDateTo = f.Dates[3], amtFrom = f.AmountFrom,
            amtTo = f.AmountTo, mthFrom = B(r.MonthFrom), mthTo = B(r.MonthTo), yrFrom = B(r.YearFrom), yrTo = B(r.YearTo),
        }, 600);
    }
}
