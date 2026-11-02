using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Invoice;

/// <summary>A row of SP_SeachCasByOPBatch.</summary>
public sealed record BatchRow
{
    /// <summary>Cases.</summary>
    public int TotNoOfCase { get; init; }
    /// <summary>Invoice number.</summary>
    public string? InvoiceNo { get; init; }
    /// <summary>Clinic name.</summary>
    public string? ClnName { get; init; }
    /// <summary>Invoice statement amount.</summary>
    public decimal? InvoiceStmtAmt { get; init; }
}

/// <summary>ShowInfo: a batch's totals.</summary>
/// <param name="BatchNo">Batch number (upper case, trimmed).</param>
/// <param name="Cases">No. of cases.</param>
/// <param name="InvoiceNo">Invoice number(s), comma-separated when the batch has several.</param>
/// <param name="InvoiceAmt">Invoice amount.</param>
/// <param name="Clinic">Clinic name (the last row's).</param>
/// <param name="Warning">"Invoice No : …, Different Clinic Found !" when the batch has more than one row.</param>
public record BatchInfo(string BatchNo, int Cases, string InvoiceNo, decimal InvoiceAmt, string Clinic, string? Warning);

/// <summary>Transfer of batch 1's cases to batch 2.</summary>
/// <param name="FromBatch">Batch No. 1.</param>
/// <param name="ToBatch">Batch No. 2.</param>
public record BatchTransferRequest(string? FromBatch, string? ToBatch);

/// <summary>Column lengths (CISCAS.OPBatchNo).</summary>
public sealed class BatchTransferRequestValidator : RequestValidator<BatchTransferRequest>
{
    /// <summary>Defines the rules.</summary>
    public BatchTransferRequestValidator()
    {
        RuleFor(r => r.FromBatch).MaximumLength(20);
        RuleFor(r => r.ToBatch).MaximumLength(20);
    }
}

/// <summary>Data access of Batch Transfer (<c>cis.BatchTransfer_*</c>).</summary>
public interface IBatchTransferRepository
{
    /// <summary>SP_SeachCasByOPBatch rows.</summary>
    Task<List<BatchRow>> BatchAsync(IDbSession session, string batchNo);

    /// <summary>BatchTransProcess; cases moved.</summary>
    Task<int> SaveAsync(IDbSession session, string fromBatch, string toBatch, string remark, decimal amount);
}

/// <summary>Invoice > Clinic > Batch Transfer.</summary>
public interface IBatchTransferService
{
    /// <summary>txtBatchNo1_LostFocus / txtBatchNo2_LostFocus → ShowInfo.</summary>
    Task<BatchInfo> BatchAsync(string? batchNo);

    /// <summary>CmdSave_Click (Transfer).</summary>
    Task<string> SaveAsync(BatchTransferRequest request, string userCode);
}

/// <inheritdoc />
public sealed class BatchTransferService(IDbSessions sessions, IBatchTransferRepository repository, TimeProvider clock) : IBatchTransferService
{
    /// <summary>Desktop texts (spacing kept).</summary>
    public const string Invalid = "Invalid OP Batch Number ! ", BatchFirst = "Please confirm your Batch Number ! ",
        DifferentInvoice = "Different Invoice No. Must be SAME Invoice No. ! ", SameBatch = "Same Batch Number ! ", Saved = "Records Transfered Successfully ! ";

    /// <inheritdoc />
    public async Task<BatchInfo> BatchAsync(string? batchNo)
    {
        var batch = Chk(batchNo);
        if (batch.Length == 0) throw new BusinessRuleException(Invalid);
        await using var s = await sessions.OpenAsync();
        return Info(batch, await repository.BatchAsync(s, batch));
    }

    /// <summary>ShowInfo's totals from the procedure's rows.</summary>
    public static BatchInfo Info(string batch, List<BatchRow> rows)
    {
        if (rows.Count == 0) throw new NotFoundException(Invalid);
        var first = rows[0].InvoiceNo?.Trim() ?? "";
        var invoice = first;
        foreach (var r in rows.Skip(1))
            if ((r.InvoiceNo?.Trim() ?? "") != first) invoice += "," + (r.InvoiceNo?.Trim() ?? "");
        // SInvAmt = CDbl(SInvAmt) + Format(amount, "#,##0.00"): each amount rounded to cents before adding.
        var amount = rows.Sum(r => Math.Round(r.InvoiceStmtAmt ?? 0, 2, MidpointRounding.AwayFromZero));
        return new BatchInfo(batch, rows.Sum(r => r.TotNoOfCase), invoice, amount, rows[^1].ClnName?.Trim() ?? "",
            rows.Count > 1 ? $"Invoice No : {invoice}, Different Clinic Found !" : null);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(BatchTransferRequest r, string userCode)
    {
        string from = Chk(r.FromBatch), to = Chk(r.ToBatch);
        if (from.Length == 0 || to.Length == 0) throw new BusinessRuleException(BatchFirst);
        await using var s = await sessions.BeginAsync();
        // The desktop compares what ShowInfo filled in on its form; the web reads both batches again.
        var one = Info(from, await repository.BatchAsync(s, from));
        var two = Info(to, await repository.BatchAsync(s, to));
        if (one.InvoiceNo != two.InvoiceNo) throw new BusinessRuleException(DifferentInvoice);
        if (from == to) throw new BusinessRuleException(SameBatch);
        var remark = $"Transfer By : {userCode} {from}->{to} , {clock.GetLocalNow():yyyyMMdd}";
        await repository.SaveAsync(s, from, to, remark, one.InvoiceAmt + two.InvoiceAmt);
        await s.CommitAsync();
        return Saved;
    }

    /// <summary>ChkStr: upper case, trimmed.</summary>
    private static string Chk(string? text) => (text ?? "").Trim().ToUpper(CultureInfo.InvariantCulture);
}
