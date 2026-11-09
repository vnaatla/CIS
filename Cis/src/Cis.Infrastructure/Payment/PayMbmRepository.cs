using Cis.Application.Abstractions;
using Cis.Application.Payment;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Payment;

/// <summary><see cref="IPayMbmRepository"/> on cis.PayMbm_*.</summary>
public sealed class PayMbmRepository : IPayMbmRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayMbmSearchRequest r, DateTime? pvDateFrom, DateTime? pvDateTo, string? pvNo,
        string? caseNo)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? M(string? v) => B(v) is { } m ? m.PadLeft(2, '0') : null;
        return session.RowsAsync("cis.PayMbm_Search", new
        {
            pvNo, caseNo, scan = r.Scan, caseFrom = B(r.CaseFrom), caseTo = r.Scan ? null : B(r.CaseTo), mbmNo = B(r.MbmNo), mbmName = B(r.MbmName), company = B(r.Company),
            payment = r.Payment, pvFrom = B(r.PvFrom), pvTo = B(r.PvTo), blankBatch = r.BlankBatch, opFrom = B(r.OpFrom), opTo = B(r.OpTo), blankPv = r.BlankPv,
            pvDateFrom, pvDateTo, blankPvDate = r.BlankPvDate, chqFrom = B(r.ChequeFrom), chqTo = B(r.ChequeTo), blankChq = r.BlankCheque, crFrom = B(r.CrFrom),
            crTo = B(r.CrTo), blankCr = r.BlankCr, mthFrom = M(r.MonthFrom), mthTo = M(r.MonthTo), blankMth = r.BlankMonth, yrFrom = B(r.YearFrom), yrTo = B(r.YearTo),
            blankYr = r.BlankYear, bordx = r.Bordx, bordxFrom = B(r.BordxFrom), bordxTo = B(r.BordxTo), btbg = r.Btbg, btbgFrom = B(r.BtbgFrom), btbgTo = B(r.BtbgTo),
        }, 600);
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ExtraAsync(IDbSession session, string cases) => session.RowsAsync("cis.PayMbm_Extra", new { cases }, 300);

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, PayMbmRow r) =>
        session.ExecAsync("cis.PayMbm_Save", new
        {
            index = r.Index, casNumber = r.CasNumber, pvNo = r.PvNo, pvDate = r.PvDate, chequeDate = r.ChequeDate, chequeNo = r.ChequeNo, payTo = r.PayTo, bank = r.Bank,
            account = r.Account, remark = r.Remark, paidAmt = r.PaidAmt, user = r.User, mth = r.Month, yr = r.Year, status = r.Status,
        });

    /// <inheritdoc />
    public Task PvChequeAsync(IDbSession session, string pvNo, DateTime chequeDate, string chequeNo, decimal amount, string payTo, string one, string two) =>
        session.ExecAsync("cis.PayMbm_PvCheque", new { pvNo, chequeDate, chequeNo, chequeAmt = amount, payTo, wordingOne = one, wordingTwo = two });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> PvRowsAsync(IDbSession session, string pvNo, string? casNumber) =>
        session.RowsAsync("cis.PayMbm_PvRows", new { pvNo, casNumber });

    /// <inheritdoc />
    public Task EditAsync(IDbSession session, string index, string casNumber, DateTime pvDate, DateTime chequeDate, string chequeNo, decimal amount, string payTo,
        string? remark, string one, string two) =>
        session.ExecAsync("cis.PayMbm_Edit", new
        {
            index, casNumber, pvDate, chequeDate, chequeNo, chequeAmt = amount, payTo, remark, wordingOne = one, wordingTwo = two,
        });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> PvPrintAsync(IDbSession session, string? cases, string? pvFrom, string? pvTo) =>
        session.RowsAsync("cis.PayMbm_PvPrint", new { cases, pvFrom, pvTo }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ChequesAsync(IDbSession session, string from, string to, bool account) =>
        session.RowsAsync("cis.PayMbm_Cheques", new { chequeFrom = from, chequeTo = to, account });

    /// <inheritdoc />
    public Task ChequesPrintedAsync(IDbSession session, string from, string to) =>
        session.ExecAsync("cis.PayMbm_ChequesPrinted", new { chequeFrom = from, chequeTo = to });
}
