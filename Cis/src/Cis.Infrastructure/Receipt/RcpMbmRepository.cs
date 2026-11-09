using Cis.Application.Abstractions;
using Cis.Application.Receipt;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Receipt;

/// <summary><see cref="IRcpMbmRepository"/> on cis.RcpMbm_*.</summary>
public sealed class RcpMbmRepository : IRcpMbmRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, RcpMbmSearchRequest r, DateTime? from, DateTime? to, decimal? amtFrom, decimal? amtTo)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        return session.RowsAsync("cis.RcpMbm_Search", new
        {
            scan = r.Scan, caseFrom = B(r.CaseFrom), caseTo = r.Scan ? null : B(r.CaseTo), mbmNo = B(r.MbmNo), mbmName = B(r.MbmName), payment = r.Payment,
            receiptFrom = B(r.ReceiptFrom), blankReceipt = r.BlankReceipt, recDateFrom = from, recDateTo = to, drFrom = B(r.DrFrom), drTo = B(r.DrTo), blankDr = r.BlankDr,
            amtFrom, amtTo, blankAmt = r.BlankAmount, mthFrom = B(r.MonthFrom), mthTo = B(r.MonthTo), blankMth = r.BlankMonth, yrFrom = B(r.YearFrom), yrTo = B(r.YearTo),
            blankYr = r.BlankYear,
        }, 300);
    }

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, string cas, string receiptNo, DateTime date, string chequeNo, string? remarks, string? mth, string? yr, string status,
        decimal amount, string user) =>
        session.ExecAsync("cis.RcpMbm_Save", new { casNumber = cas, receiptNo, receiptDate = date, chequeNo, remarks, mth, yr, status, amount, user });

    /// <inheritdoc />
    public Task EditAsync(IDbSession session, string? index, string receiptNo, DateTime date, decimal amount, string chequeNo, string? remarks, string? mth, string? yr,
        string user) =>
        session.ExecAsync("cis.RcpMbm_Edit", new { index, receiptNo, receiptDate = date, amount, chequeNo, remarks, mth, yr, user });

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> ReceiptAsync(IDbSession session, string index) =>
        (await session.RowsAsync("cis.RcpMbm_Receipt", new { index })).FirstOrDefault();
}
