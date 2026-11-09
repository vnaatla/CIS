using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Receipt;

/// <summary>Receipt from FFS search (FrmReceiptFFS SearchRecord).</summary>
public sealed record RcpFfsSearchRequest
{
#pragma warning disable CS1591
    public bool Scan { get; init; } = true;
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    /// <summary>0 All, 1 Paid, 2 Unpaid.</summary>
    public int Payment { get; init; }
    public string? ReceiptFrom { get; init; }
    public string? ReceiptTo { get; init; }
    public string? ReceiptDateFrom { get; init; }
    public string? ReceiptDateTo { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
    public string? BordxDateFrom { get; init; }
    public string? BordxDateTo { get; init; }
    public string? Company { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpFfsSearchRequestValidator : Common.RequestValidator<RcpFfsSearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpFfsSearchRequestValidator()
    {
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.Company).MaximumLength(200);
    }
}

/// <summary>Save (the ticked bordereaux) / Update (the selected receipt).</summary>
public sealed record RcpFfsSaveRequest
{
#pragma warning disable CS1591
    public string? Index { get; init; }
    public string? BordxNo { get; init; }
    public string? ReceiptNo { get; init; }
    public string? ReceiptDate { get; init; }
    public string? ReceiptAmt { get; init; }
    public string? ChequeNo { get; init; }
    public string? Remark { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    public List<string> Ticked { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpFfsSaveRequestValidator : Common.RequestValidator<RcpFfsSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpFfsSaveRequestValidator()
    {
        RuleFor(r => r.ReceiptNo).MaximumLength(30);
        RuleFor(r => r.ChequeNo).MaximumLength(30);
        RuleFor(r => r.Remark).MaximumLength(500);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.Length <= 20));
    }
}

/// <summary>The listing and the total FFS.</summary>
public record RcpFfsResult(List<Dictionary<string, object?>> Rows, decimal TotalFfs);

/// <summary>Data access.</summary>
public interface IRcpFfsRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, RcpFfsSearchRequest r, DateTime?[] dates, string? receiptNo);
    Task SaveAsync(IDbSession session, string? index, string bordxNo, string receiptNo, DateTime date, string chequeNo, decimal amount, string? remarks, string mth,
        string yr, string user);
#pragma warning restore CS1591
}

/// <summary>Receipt > Fr FFS (FrmReceiptFFS).</summary>
public interface IRcpFfsService
{
    /// <summary>CmdSearch_Click (or the receipt just saved).</summary>
    Task<RcpFfsResult> SearchAsync(RcpFfsSearchRequest request, string? receiptNo = null);

    /// <summary>CmdSave_Click (after the confirmation).</summary>
    Task<string> SaveAsync(RcpFfsSaveRequest request, string user);

    /// <summary>CmdUpdate_Click (after the confirmation).</summary>
    Task<string> UpdateAsync(RcpFfsSaveRequest request, string user);
}

/// <inheritdoc />
public sealed class RcpFfsService(IDbSessions sessions, IRcpFfsRepository repository) : IRcpFfsService
{
    /// <summary>Desktop texts.</summary>
    public const string NoBordx = "Please Enter Bordx No", NoRecord = "No record found!", NoSelection = "Please select the record", NoReceiptNo = "Please Enter Receipt No",
        BadDate = "Invalid Receipt Date!", NoAmount = "Please Enter Receipt Amt!", WrongAmount = "Amount to paid is incorrect\nPlease key in remarks.",
        NoCheque = "Please Cheque Number!", NoMonth = "Please keyin Account Period Month ! ", NoYear = "Please keyin Account Period Year ! ", Saved = "Record Saved...",
        Received = "FFS has been received!";

    /// <inheritdoc />
    public async Task<RcpFfsResult> SearchAsync(RcpFfsSearchRequest r, string? receiptNo = null)
    {
        if (receiptNo is null && r.Scan && string.IsNullOrWhiteSpace(r.BordxFrom)) throw new BusinessRuleException(NoBordx);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r, [ClinicInvoiceRules.Date(r.ReceiptDateFrom), ClinicInvoiceRules.Date(r.ReceiptDateTo),
            ClinicInvoiceRules.Date(r.BordxDateFrom), ClinicInvoiceRules.Date(r.BordxDateTo)], receiptNo);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        foreach (var x in rows) Shape(x);
        return new RcpFfsResult(rows, rows.Sum(x => N(x, "CLMMCOProcFee")));
    }

    /// <summary>ReceiptBordxListing / LstFFSRcpt_ItemCheck: status, balance (FFS − received) and the tick message.</summary>
    public static void Shape(Dictionary<string, object?> x)
    {
        var status = T(x, "ReceiptStatus");
        x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
        x["Balance"] = N(x, "CLMMCOProcFee") - N(x, "ReceiptAmt");
        x["TickError"] = status == "F" ? Received : null;
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(RcpFfsSaveRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoSelection);
        var (date, cheque) = Check(r);
        await using var tx = await sessions.BeginAsync();
        var rows = new List<Dictionary<string, object?>>();
        foreach (var b in r.Ticked.Select(x => x.Trim()).Distinct())
        {
            var row = (await repository.SearchAsync(tx, new RcpFfsSearchRequest { BordxFrom = b }, new DateTime?[4], null)).FirstOrDefault()
                ?? throw new BusinessRuleException(NoSelection);
            Shape(row);
            if (row["TickError"] is string e) throw new BusinessRuleException(e);
            rows.Add(row);
        }
        if (Amount(r.ReceiptAmt) != rows.Sum(x => N(x, "CLMMCOProcFee")) && string.IsNullOrWhiteSpace(r.Remark)) throw new BusinessRuleException(WrongAmount);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        foreach (var x in rows)
            await repository.SaveAsync(tx, null, T(x, "BordxRefNo"), r.ReceiptNo!.Trim(), date, cheque, N(x, "CLMMCOProcFee"), Blank(r.Remark), r.Month.Trim(), r.Year.Trim(),
                user);
        await tx.CommitAsync();
        return Saved;
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(RcpFfsSaveRequest r, string user)
    {
        var (date, cheque) = Check(r);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        if (string.IsNullOrWhiteSpace(r.Index) || string.IsNullOrWhiteSpace(r.BordxNo)) throw new BusinessRuleException(NoSelection);
        await using var tx = await sessions.BeginAsync();
        await repository.SaveAsync(tx, r.Index.Trim(), r.BordxNo.Trim(), r.ReceiptNo!.Trim(), date, cheque, Amount(r.ReceiptAmt), Blank(r.Remark), r.Month.Trim(),
            r.Year.Trim(), user);
        await tx.CommitAsync();
        return Saved;
    }

    private static (DateTime Date, string Cheque) Check(RcpFfsSaveRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(NoReceiptNo);
        var date = ClinicInvoiceRules.Date(r.ReceiptDate) ?? throw new BusinessRuleException(BadDate);
        if (string.IsNullOrWhiteSpace(r.ReceiptAmt)) throw new BusinessRuleException(NoAmount);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        return (date, r.ChequeNo.Trim());
    }

    private static decimal Amount(string? v) => decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
