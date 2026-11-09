using Cis.Application.Abstractions;
using Cis.Application.Receipt;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Receipt;

/// <summary><see cref="IRcpPayRepository"/> on cis.RcpPay_* (and HISDB SearchSubsidiary).</summary>
public sealed class RcpPayRepository : IRcpPayRepository
{
    /// <inheritdoc />
    public async Task<(List<string> Companies, List<(string Code, string Name)> Payors)> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.RcpPay_Lookups");
        var companies = (await grid.ReadAsync<string>()).ToList();
        var payors = (await grid.ReadAsync<(string Code, string Name)>()).ToList();
        return (companies, payors);
    }

    /// <inheritdoc />
    public Task<List<string>> ClientsAsync(IDbSession session, string payor) => session.QueryAsync<string>("cis.RcpPay_Clients", new { payor });

    /// <inheritdoc />
    public async Task<List<string>> SubsidiariesAsync(IDbSession session, string company) =>
        (await session.RowsAsync("HISDB.dbo.SearchSubsidiary", new { GroupCompany = company }))
            .Select(x => Convert.ToString(x.GetValueOrDefault("MBMGRPSubsidiary"))?.Trim() ?? "").Where(x => x.Length > 0).Distinct().ToList();

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, RcpPaySearchRequest r, string? payor, DateTime? from, DateTime? to, string? receiptNo)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        return session.RowsAsync("cis.RcpPay_Search", new
        {
            worksheet = r.Worksheet, scan = r.Scan, receiptNo, company = B(r.Company), subsidiary = B(r.Subsidiary), bordxFrom = B(r.BordxFrom), bordxTo = B(r.BordxTo),
            caseFrom = B(r.CaseFrom), caseTo = B(r.CaseTo), wsStatus = r.WsStatus, bordxStatus = r.BordxStatus, recFrom = B(r.ReceiptFrom), recTo = B(r.ReceiptTo),
            recDateFrom = from, recDateTo = to, mthFrom = B(r.MonthFrom), mthTo = B(r.MonthTo), yrFrom = B(r.YearFrom), yrTo = B(r.YearTo), payor,
            client = B(r.InsuredClient), pvFrom = B(r.ClientPvFrom), pvTo = B(r.ClientPvTo), mrFrom = B(r.MedixReceiptFrom), mrTo = B(r.MedixReceiptTo),
            insured = B(r.Insured),
        }, 600);
    }

    /// <inheritdoc />
    public async Task<(string Status, List<string> Cases)> BordxCasesAsync(IDbSession session, string bordxNo)
    {
        using var grid = await session.MultipleAsync("cis.RcpPay_BordxCases", new { bordxNo });
        var status = (await grid.ReadFirstOrDefaultAsync<string>()) ?? "";
        return (status.Trim(), (await grid.ReadAsync<string>()).ToList());
    }

    /// <inheritdoc />
    public async Task<string> CaseStatusAsync(IDbSession session, string casNumber, string bordxNo) =>
        (await session.FirstOrDefaultAsync<string>("cis.RcpPay_CaseStatus", new { casNumber, bordxNo }))?.Trim() ?? "";

    /// <inheritdoc />
    public Task SaveBordxCaseAsync(IDbSession session, string casNumber, string bordxNo, RcpPayReceipt r) =>
        session.ExecAsync("cis.RcpPay_SaveBordxCase", new
        {
            casNumber, bordxNo, receiptNo = r.ReceiptNo, receiptDate = r.ReceiptDate, receiptAmt = r.ReceiptAmt, chequeNo = r.ChequeNo, remark = r.Remark, mth = r.Month,
            yr = r.Year, user = r.User,
        });

    /// <inheritdoc />
    public Task SaveWksCaseAsync(IDbSession session, string casNumber, string bordxNo, string status, RcpPayReceipt r) =>
        session.ExecAsync("cis.RcpPay_SaveWksCase", new
        {
            casNumber, bordxNo, status, receiptNo = r.ReceiptNo, receiptDate = r.ReceiptDate, receiptAmt = r.ReceiptAmt, chequeNo = r.ChequeNo, remark = r.Remark,
            mth = r.Month, yr = r.Year, user = r.User,
        });

    /// <inheritdoc />
    public Task SummaryAsync(IDbSession session, string receiptNo, DateTime date, decimal amount, string type, string? desc, string chequeNo, string payCode) =>
        session.ExecAsync("cis.RcpPay_Summary", new { receiptNo, receiptDate = date, amount, type, desc, chequeNo, payCode });

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> ReceiptAsync(IDbSession session, string bordxNo, string? casNumber) =>
        (await session.RowsAsync("cis.RcpPay_Receipt", new { bordxNo, casNumber })).FirstOrDefault();

    /// <inheritdoc />
    public Task EditAsync(IDbSession session, string bordxNo, string? casNumber, RcpPayReceipt r) =>
        session.ExecAsync("cis.RcpPay_Edit", new
        {
            bordxNo, casNumber, receiptNo = r.ReceiptNo, receiptDate = r.ReceiptDate, receiptAmt = r.ReceiptAmt, chequeNo = r.ChequeNo, remark = r.Remark, mth = r.Month,
            yr = r.Year, user = r.User,
        });
}
