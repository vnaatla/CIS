using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Receipt;

/// <summary>Receipt from Member search (FrmReceiptMember SearchRecord).</summary>
public sealed record RcpMbmSearchRequest
{
#pragma warning disable CS1591
    public bool Scan { get; init; } = true;
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    public string? MbmNo { get; init; }
    public string? MbmName { get; init; }
    /// <summary>0 All, 1 Paid, 2 Unpaid, 3 Partial, 4 Paid &amp; Partial, 5 Unpaid &amp; Partial.</summary>
    public int Payment { get; init; }
    public string? ReceiptFrom { get; init; }
    public bool BlankReceipt { get; init; }
    public string? ReceiptDateFrom { get; init; }
    public string? ReceiptDateTo { get; init; }
    public string? DrFrom { get; init; }
    public string? DrTo { get; init; }
    public bool BlankDr { get; init; }
    public string? AmountFrom { get; init; }
    public string? AmountTo { get; init; }
    public bool BlankAmount { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    public bool BlankMonth { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
    public bool BlankYear { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpMbmSearchRequestValidator : Common.RequestValidator<RcpMbmSearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpMbmSearchRequestValidator()
    {
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
        RuleFor(r => r.MbmNo).MaximumLength(30);
        RuleFor(r => r.MbmName).MaximumLength(200);
    }
}

/// <summary>Save (the ticked cases) / Update (the selected receipt).</summary>
public sealed record RcpMbmSaveRequest
{
#pragma warning disable CS1591
    public string? Index { get; init; }
    public string? ReceiptNo { get; init; }
    public string? ReceiptDate { get; init; }
    public string? Amount { get; init; }
    public string? ChequeNo { get; init; }
    public string? Remarks { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    /// <summary>The ticked case numbers.</summary>
    public List<string> Ticked { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpMbmSaveRequestValidator : Common.RequestValidator<RcpMbmSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpMbmSaveRequestValidator()
    {
        RuleFor(r => r.ReceiptNo).MaximumLength(30);
        RuleFor(r => r.ChequeNo).MaximumLength(30);
        RuleFor(r => r.Remarks).MaximumLength(500);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.Length <= 30));
    }
}

/// <summary>Data access.</summary>
public interface IRcpMbmRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, RcpMbmSearchRequest r, DateTime? from, DateTime? to, decimal? amtFrom, decimal? amtTo);
    Task SaveAsync(IDbSession session, string cas, string receiptNo, DateTime date, string chequeNo, string? remarks, string? mth, string? yr, string status, decimal amount,
        string user);
    Task EditAsync(IDbSession session, string? index, string receiptNo, DateTime date, decimal amount, string chequeNo, string? remarks, string? mth, string? yr, string user);
    Task<Dictionary<string, object?>?> ReceiptAsync(IDbSession session, string index);
#pragma warning restore CS1591
}

/// <summary>Receipt > Fr Member (FrmReceiptMember).</summary>
public interface IRcpMbmService
{
    /// <summary>cmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(RcpMbmSearchRequest request);

    /// <summary>lstReceiptMBM_Click in edit mode: the receipt.</summary>
    Task<Dictionary<string, object?>> ReceiptAsync(string? index);

    /// <summary>cmdSave_Click (after the confirmation).</summary>
    Task<string> SaveAsync(RcpMbmSaveRequest request, string user);

    /// <summary>cmdUpdate_Click (after the confirmation).</summary>
    Task<string> UpdateAsync(RcpMbmSaveRequest request, string user);
}

/// <inheritdoc />
public sealed class RcpMbmService(IDbSessions sessions, IRcpMbmRepository repository) : IRcpMbmService
{
    /// <summary>Desktop texts.</summary>
    public const string NoWorksheet = "Please Enter Worksheet No", NoRecord = "No record found!", NoReceiptNo = "Please Enter Receipt No",
        NoReceiptDate = "Please Enter Receipt Date", NoAmount = "Please Enter Amount", NoCheque = "Please Enter Cheque No", Updated = "Record updated...",
        NoSelection = "Please select the record", NoReceiptNoUpd = "Please Enter Receipt No!", BadDate = "Invalid Receipt Date!", NoChequeUpd = "Please Enter Cheque No!",
        NoAmountUpd = "Please Enter Cheque Amt!", UpdatedEdit = "Record Updated...", Paid = "Payment has been made!", NotPaid = "Payment hasn't been made!";

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(RcpMbmSearchRequest r)
    {
        if (r.Scan && string.IsNullOrWhiteSpace(r.CaseFrom)) throw new BusinessRuleException(NoWorksheet);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r, ClinicInvoiceRules.Date(r.ReceiptDateFrom), ClinicInvoiceRules.Date(r.ReceiptDateTo), Amount(r.AmountFrom),
            Amount(r.AmountTo));
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        foreach (var x in rows) Shape(x);
        return rows;
    }

    /// <summary>ReceiptFrMBMListing / ProcessCheckedItem: status, the debit amount (shown positive), received, balance, amount to receive, tick message.</summary>
    public static void Shape(Dictionary<string, object?> x)
    {
        var status = T(x, "PVRecMBMStatus");
        var amount = -N(x, "CLMPayableByMedix2M");
        var received = status == "P" ? N(x, "ReceivedTotal") : N(x, "ReceiptAmt");
        x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
        x["Amount"] = amount;
        x["Received"] = status.Length > 0 ? received : null;
        x["Balance"] = status switch { "P" => amount - received, "" => amount, _ => null };
        x["AmountToReceive"] = status == "P" ? amount - received : amount;
        x["TickError"] = status == "F" ? Paid : null;
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>> ReceiptAsync(string? index)
    {
        await using var s = await sessions.OpenAsync();
        return (string.IsNullOrWhiteSpace(index) ? null : await repository.ReceiptAsync(s, index.Trim())) ?? throw new BusinessRuleException(NotPaid);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(RcpMbmSaveRequest r, string user)
    {
        if (string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(NoReceiptNo);
        var date = ClinicInvoiceRules.Date(r.ReceiptDate) ?? throw new BusinessRuleException(NoReceiptDate);
        if (string.IsNullOrWhiteSpace(r.Amount)) throw new BusinessRuleException(NoAmount);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        await using var tx = await sessions.BeginAsync();
        foreach (var cas in r.Ticked.Select(c => c.Trim()).Where(c => c.Length > 0).Distinct())
        {
            var row = (await repository.SearchAsync(tx, new RcpMbmSearchRequest { Scan = true, CaseFrom = cas }, null, null, null, null)).FirstOrDefault();
            if (row is null) continue;
            Shape(row);
            if (row["TickError"] is string e) throw new BusinessRuleException(e);
            await repository.SaveAsync(tx, cas, r.ReceiptNo.Trim(), date, r.ChequeNo.Trim(), Blank(r.Remarks), Blank(r.Month), Blank(r.Year), "Full",
                (decimal)row["AmountToReceive"]!, user);
        }
        await tx.CommitAsync();
        return Updated;
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(RcpMbmSaveRequest r, string user)
    {
        if (string.IsNullOrWhiteSpace(r.Index)) throw new BusinessRuleException(NoSelection);
        if (string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(NoReceiptNoUpd);
        var date = ClinicInvoiceRules.Date(r.ReceiptDate) ?? throw new BusinessRuleException(string.IsNullOrWhiteSpace(r.ReceiptDate) ? NoReceiptDate : BadDate);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoChequeUpd);
        if (string.IsNullOrWhiteSpace(r.Amount)) throw new BusinessRuleException(NoAmountUpd);
        await using var tx = await sessions.BeginAsync();
        await repository.EditAsync(tx, r.Index.Trim(), r.ReceiptNo.Trim(), date, Amount(r.Amount) ?? 0, r.ChequeNo.Trim(), Blank(r.Remarks), Blank(r.Month), Blank(r.Year), user);
        await tx.CommitAsync();
        return UpdatedEdit;
    }

    private static decimal? Amount(string? v) =>
        decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
