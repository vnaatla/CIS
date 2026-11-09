using Cis.Application.Abstractions;
using Cis.Application.Receipt;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Receipt;

/// <summary><see cref="IRcpMnrbRepository"/> on cis.RcpMnrb_*.</summary>
public sealed class RcpMnrbRepository : IRcpMnrbRepository
{
    /// <inheritdoc />
    public Task<List<(string Code, string Name)>> PayorsAsync(IDbSession session) => session.QueryAsync<(string Code, string Name)>("cis.RcpMnrb_Payors");

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, RcpMnrbSearchRequest r, string? payor, DateTime? from, DateTime? to)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        return session.RowsAsync("cis.RcpMnrb_Search", new
        {
            worksheet = r.Worksheet, scan = r.Scan, payor, bordxFrom = B(r.BordxFrom), bordxTo = B(r.BordxTo), caseFrom = B(r.CaseFrom), caseTo = B(r.CaseTo),
            wsStatus = r.WsStatus, bordxStatus = r.BordxStatus, recFrom = B(r.ReceiptFrom), recTo = B(r.ReceiptTo), recDateFrom = from, recDateTo = to,
            mthFrom = B(r.MonthFrom), mthTo = B(r.MonthTo), yrFrom = B(r.YearFrom), yrTo = B(r.YearTo),
        }, 600);
    }

    /// <inheritdoc />
    public async Task<(string Status, List<string> Cases)> BordxCasesAsync(IDbSession session, string bordxNo)
    {
        using var grid = await session.MultipleAsync("cis.RcpMnrb_BordxCases", new { bordxNo });
        var status = (await grid.ReadFirstOrDefaultAsync<string>()) ?? "";
        return (status.Trim(), (await grid.ReadAsync<string>()).ToList());
    }

    /// <inheritdoc />
    public async Task<string> CaseStatusAsync(IDbSession session, string casNumber, string bordxNo) =>
        (await session.FirstOrDefaultAsync<string>("cis.RcpMnrb_CaseStatus", new { casNumber, bordxNo }))?.Trim() ?? "";

    /// <inheritdoc />
    public Task SaveBordxCaseAsync(IDbSession session, string casNumber, string bordxNo, RcpPayReceipt r) =>
        session.ExecAsync("cis.RcpMnrb_SaveBordxCase", new
        {
            casNumber, bordxNo, receiptNo = r.ReceiptNo, receiptDate = r.ReceiptDate, receiptAmt = r.ReceiptAmt, chequeNo = r.ChequeNo, remark = r.Remark, mth = r.Month,
            yr = r.Year, user = r.User,
        });

    /// <inheritdoc />
    public Task SaveWksCaseAsync(IDbSession session, string casNumber, string bordxNo, string status, decimal selectedAmt, RcpPayReceipt r) =>
        session.ExecAsync("cis.RcpMnrb_SaveWksCase", new
        {
            casNumber, bordxNo, status, receiptNo = r.ReceiptNo, receiptDate = r.ReceiptDate, selectedAmt, receiptAmt = r.ReceiptAmt, chequeNo = r.ChequeNo,
            remark = r.Remark, mth = r.Month, yr = r.Year, user = r.User,
        });

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> ReceiptAsync(IDbSession session, string bordxNo, string? casNumber) =>
        (await session.RowsAsync("cis.RcpMnrb_Receipt", new { bordxNo, casNumber })).FirstOrDefault();

    /// <inheritdoc />
    public Task EditAsync(IDbSession session, string bordxNo, string? casNumber, RcpPayReceipt r) =>
        session.ExecAsync("cis.RcpMnrb_Edit", new
        {
            bordxNo, casNumber, receiptNo = r.ReceiptNo, receiptDate = r.ReceiptDate, receiptAmt = r.ReceiptAmt, chequeNo = r.ChequeNo, remark = r.Remark, mth = r.Month,
            yr = r.Year, user = r.User,
        });
}
