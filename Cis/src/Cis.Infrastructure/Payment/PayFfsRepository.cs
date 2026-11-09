using Cis.Application.Abstractions;
using Cis.Application.Payment;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Payment;

/// <summary><see cref="IPayFfsRepository"/> on cis.PayFfs_*.</summary>
public sealed class PayFfsRepository : IPayFfsRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayFfsSearchRequest r, string? pvNo, DateTime?[] dates)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? M(string? v) => B(v) is { } m ? m.PadLeft(2, '0') : null;
        return session.RowsAsync("cis.PayFfs_Search", new
        {
            pvNo, company = B(r.Company), payment = r.Payment, scan = r.Scan, bordxFrom = B(r.BordxFrom), bordxTo = B(r.BordxTo), blankBordxDate = r.BlankBordxDate,
            bordxDateFrom = dates[0], bordxDateTo = dates[1], blankPv = r.BlankPv, pvFrom = B(r.PvFrom), pvTo = B(r.PvTo), blankPvDate = r.BlankPvDate,
            pvDateFrom = dates[2], pvDateTo = dates[3], blankMth = r.BlankMonth, mthFrom = M(r.MonthFrom), mthTo = M(r.MonthTo), blankYr = r.BlankYear,
            yrFrom = B(r.YearFrom), yrTo = B(r.YearTo),
        }, 300);
    }

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, PayFfsRow r) =>
        session.ExecAsync("cis.PayFfs_Save", new
        {
            index = r.Index, bordxNo = r.BordxNo, pvNo = r.PvNo, pvDate = r.PvDate, chequeDate = r.ChequeDate, chequeNo = r.ChequeNo, chequeAmt = r.ChequeAmt, bank = r.Bank,
            account = r.Account, remark = r.Remark, paidAmt = r.PaidAmt, mth = r.Month, yr = r.Year, payee = r.Payee, user = r.User,
        });
}
