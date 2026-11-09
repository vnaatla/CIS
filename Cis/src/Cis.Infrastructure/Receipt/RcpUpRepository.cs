using Cis.Application.Abstractions;
using Cis.Application.Receipt;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Receipt;

/// <summary><see cref="IRcpUpRepository"/> on cis.RcpUp_*.</summary>
public sealed class RcpUpRepository : IRcpUpRepository
{
    /// <inheritdoc />
    public Task<List<(string CasNumber, string? BordxRefNo, decimal CLMBordxAmt)>> ClaimsAsync(IDbSession session, string cases) =>
        session.QueryAsync<(string CasNumber, string? BordxRefNo, decimal CLMBordxAmt)>("cis.RcpUp_Claims", new { cases });

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, string casNumber, string? bordxNo, string receiptNo, DateTime? date, decimal receiptAmt, string? chequeNo,
        decimal excelAmt, string? mth, string? yr, string? clientPvNo, string? medixReceiptNo, string user) =>
        session.ExecAsync("cis.RcpUp_Save", new
        {
            casNumber, bordxNo, receiptNo, receiptDate = date, receiptAmt, chequeNo, excelAmt, mth, yr, clientPvNo, medixReceiptNo, user,
        });
}
