using Cis.Application.Abstractions;
using Cis.Application.Payment;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Payment;

/// <summary><see cref="IPayMbmOldRepository"/> on cis.PayMbmOld_*.</summary>
public sealed class PayMbmOldRepository : IPayMbmOldRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayMbmSearchRequest r, DateTime? pvDateFrom, DateTime? pvDateTo, string? pvNo)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? M(string? v) => B(v) is { } m ? m.PadLeft(2, '0') : null;
        return session.RowsAsync("cis.PayMbmOld_Search", new
        {
            pvNo, scan = r.Scan, caseFrom = B(r.CaseFrom), caseTo = r.Scan ? null : B(r.CaseTo), mbmNo = B(r.MbmNo), mbmName = B(r.MbmName), company = B(r.Company),
            payment = r.Payment, pvFrom = B(r.PvFrom), pvTo = B(r.PvTo), blankBatch = r.BlankBatch, opFrom = B(r.OpFrom), opTo = B(r.OpTo), blankPv = r.BlankPv,
            pvDateFrom, pvDateTo, blankPvDate = r.BlankPvDate, chqFrom = B(r.ChequeFrom), chqTo = B(r.ChequeTo), blankChq = r.BlankCheque, crFrom = B(r.CrFrom),
            crTo = B(r.CrTo), blankCr = r.BlankCr, mthFrom = M(r.MonthFrom), mthTo = M(r.MonthTo), blankMth = r.BlankMonth, yrFrom = B(r.YearFrom), yrTo = B(r.YearTo),
            blankYr = r.BlankYear, bordx = r.Bordx,
        }, 300);
    }

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, PayMbmOldRow r) =>
        session.ExecAsync("cis.PayMbmOld_Save", new
        {
            edit = r.Edit, index = r.Index, casNumber = r.CasNumber, pvNo = r.PvNo, pvDate = r.PvDate, chequeDate = r.ChequeDate, chequeNo = r.ChequeNo,
            chequeAmt = r.ChequeAmt, payTo = r.PayTo, bank = r.Bank, account = r.Account, remark = r.Remark, paidAmt = r.PaidAmt, user = r.User, mth = r.Month,
            yr = r.Year, status = r.Status, wordingOne = r.WordingOne, wordingTwo = r.WordingTwo,
        });
}
