using Cis.Application.Abstractions;
using Cis.Application.Receipt;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Receipt;

/// <summary><see cref="IRcpFfsRepository"/> on cis.RcpFfs_*.</summary>
public sealed class RcpFfsRepository : IRcpFfsRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, RcpFfsSearchRequest r, DateTime?[] dates, string? receiptNo)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? M(string? v) => B(v) is { } m ? m.PadLeft(2, '0') : null;
        return session.RowsAsync("cis.RcpFfs_Search", new
        {
            receiptNo, scan = r.Scan, bordxFrom = B(r.BordxFrom), bordxTo = r.Scan ? null : B(r.BordxTo), payment = r.Payment, recFrom = B(r.ReceiptFrom),
            recTo = B(r.ReceiptTo), recDateFrom = dates[0], recDateTo = dates[1], mthFrom = M(r.MonthFrom), mthTo = M(r.MonthTo), yrFrom = B(r.YearFrom), yrTo = B(r.YearTo),
            bordxDateFrom = dates[2], bordxDateTo = dates[3], company = B(r.Company),
        }, 300);
    }

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, string? index, string bordxNo, string receiptNo, DateTime date, string chequeNo, decimal amount, string? remarks,
        string mth, string yr, string user) =>
        session.ExecAsync("cis.RcpFfs_Save", new { index, bordxNo, receiptNo, receiptDate = date, chequeNo, amount, remarks, mth, yr, user });
}
