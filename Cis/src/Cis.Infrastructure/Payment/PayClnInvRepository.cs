using Cis.Application.Abstractions;
using Cis.Application.Payment;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Payment;

/// <summary><see cref="IPayClnInvRepository"/> on cis.PayClnInv_*.</summary>
public sealed class PayClnInvRepository : IPayClnInvRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayClnInvSearchRequest r, string? clinic, DateTime? pvDateFrom, DateTime? pvDateTo,
        string? pvNo, string? pvLast)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? M(string? v) => B(v) is { } m ? m.PadLeft(2, '0') : null;
        return session.RowsAsync("cis.PayClnInv_Search", new
        {
            pvNo, pvLast, clinic, yrFrom = B(r.YearFrom), yrTo = B(r.YearTo), mthFrom = M(r.MonthFrom), mthTo = M(r.MonthTo), payment = r.Payment, pvFrom = B(r.PvFrom),
            pvTo = B(r.PvTo), blankPv = r.BlankPv, pvDateFrom, pvDateTo, blankPvDate = r.BlankPvDate, acMthFrom = M(r.AcMonthFrom), acMthTo = M(r.AcMonthTo),
            blankAcMth = r.BlankAcMonth, acYrFrom = B(r.AcYearFrom), acYrTo = B(r.AcYearTo), blankAcYr = r.BlankAcYear, entity = r.Entity,
        }, 600);
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ClinicsAsync(IDbSession session, string codes) => session.RowsAsync("cis.PayClnInv_Clinics", new { codes });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> InvoicesAsync(IDbSession session, string? clinic, string? year, string? month, string? pvFrom, string? pvTo) =>
        session.RowsAsync("cis.PayClnInv_Invoices", new { clinic, yr = year, mth = month?.PadLeft(2, '0'), pvFrom, pvTo }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> CasesAsync(IDbSession session, string clinic, string year, string months, int entity) =>
        session.RowsAsync("cis.PayClnInv_Cases", new { clinic, yr = year, months, entity }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> PvRowsAsync(IDbSession session, string pvNo) => session.RowsAsync("cis.PayClnInv_PvRows", new { pvNo });

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, PayClnInvRow r) =>
        session.ExecAsync("cis.PayClnInv_Save", new
        {
            index = r.Index, casNumber = r.CasNumber, pvNo = r.PvNo, pvDate = r.PvDate, chequeDate = r.ChequeDate, chequeNo = r.ChequeNo, chequeAmt = r.ChequeAmt,
            bank = r.Bank, remark = r.Remark, paidAmt = r.PaidAmt, mth = r.Month, yr = r.Year, status = r.Status, wordingOne = r.WordingOne, wordingTwo = r.WordingTwo,
            user = r.User,
        });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ChequesAsync(IDbSession session, string from, string to) =>
        session.RowsAsync("cis.PayClnInv_Cheques", new { chequeFrom = from, chequeTo = to });

    /// <inheritdoc />
    public Task ChequesPrintedAsync(IDbSession session, string from, string to) =>
        session.ExecAsync("cis.PayClnInv_ChequesPrinted", new { chequeFrom = from, chequeTo = to });
}
