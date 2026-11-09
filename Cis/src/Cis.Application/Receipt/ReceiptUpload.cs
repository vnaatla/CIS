using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Receipt;

/// <summary>An uploaded row as compared with CIS (lstReceiptListing).</summary>
public sealed record RcpUpRow(string ClaimNo, string BordxNo, decimal Amount, string? ClaimNoCis, string? BordxNoCis, decimal AmountCis, string CaseStatus,
    string BordxStatus, decimal Difference, string? ClientPvNo, string? ReceiptNo, string? MedixReceiptNo);

/// <summary>Update: the receipt fields and the uploaded rows.</summary>
public sealed record RcpUpSaveRequest
{
#pragma warning disable CS1591
    public bool Multiple { get; init; }
    public string? ReceiptNo { get; init; }
    public string? MedixReceiptNo { get; init; }
    public string? ReceiptDate { get; init; }
    public string? ReceiptAmt { get; init; }
    public string? ChequeNo { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    public List<RcpUpRow> Rows { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpUpSaveRequestValidator : Common.RequestValidator<RcpUpSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpUpSaveRequestValidator()
    {
        RuleFor(r => r.ReceiptNo).MaximumLength(30);
        RuleFor(r => r.MedixReceiptNo).MaximumLength(30);
        RuleFor(r => r.ChequeNo).MaximumLength(30);
        RuleFor(r => r.Rows).Must(t => t.Count <= 60000 && t.All(x => x.ClaimNo.Length <= 30 && (x.ReceiptNo ?? "").Length <= 30 && (x.ClientPvNo ?? "").Length <= 50));
    }
}

/// <summary>Data access.</summary>
public interface IRcpUpRepository
{
#pragma warning disable CS1591
    Task<List<(string CasNumber, string? BordxRefNo, decimal CLMBordxAmt)>> ClaimsAsync(IDbSession session, string cases);
    Task SaveAsync(IDbSession session, string casNumber, string? bordxNo, string receiptNo, DateTime? date, decimal receiptAmt, string? chequeNo, decimal excelAmt,
        string? mth, string? yr, string? clientPvNo, string? medixReceiptNo, string user);
#pragma warning restore CS1591
}

/// <summary>Receipt > Upload Receipt (frmReceiptUpload) and Upload Receipt (Multiple Receipts) (frmReceiptUpMultiple).</summary>
public interface IRcpUpService
{
    /// <summary>cmpUpload_Click: the uploaded rows compared with CIS.</summary>
    Task<List<RcpUpRow>> UploadAsync(Stream content, bool multiple);

    /// <summary>CmdUpdate_Click (after the confirmation).</summary>
    Task<string> SaveAsync(RcpUpSaveRequest request, string user);
}

/// <inheritdoc />
public sealed class RcpUpService(IDbSessions sessions, IRcpUpRepository repository, ISpreadsheetReader reader) : IRcpUpService
{
    /// <summary>Desktop texts.</summary>
    public const string NoFile = "Please select file to upload!", NotUploaded = "Please Upload file..! ", NoReceiptNo = "Please Enter Receipt Number! ",
        NoMedixNo = "Please Enter Medix Receipt Number! ", Updated = "Record Updated...", Matched = "Matched", NotMatched = "Not Matched";

    /// <inheritdoc />
    public async Task<List<RcpUpRow>> UploadAsync(Stream content, bool multiple)
    {
        List<(int Row, string[] Cells)> sheet;
        try
        {
            sheet = reader.Read(content, 9, 60000, multiple ? 2 : 3, true);
        }
        catch (Exception e) when (e is not BusinessRuleException)
        {
            throw new BusinessRuleException(NoFile);
        }
        await using var s = await sessions.OpenAsync();
        var claims = new Dictionary<string, (string CasNumber, string? BordxRefNo, decimal CLMBordxAmt)>(StringComparer.OrdinalIgnoreCase);
        foreach (var chunk in sheet.Select(x => x.Cells[3].Trim()).Where(x => x.Length > 0).Distinct().Chunk(500))
            foreach (var c in await repository.ClaimsAsync(s, string.Join(',', chunk))) claims.TryAdd(c.CasNumber, c);
        return sheet.Select(x =>
        {
            var claimNo = x.Cells[3].Trim();
            var amount = Math.Round(Amount(x.Cells[6]), 2);
            var found = claims.TryGetValue(claimNo, out var c);
            return new RcpUpRow(claimNo, x.Cells[2].Trim(), amount, found ? c.CasNumber : null, found ? c.BordxRefNo : null, found ? c.CLMBordxAmt : 0,
                found ? Matched : NotMatched, found ? Matched : NotMatched, found ? c.CLMBordxAmt - amount : amount, Blank(x.Cells[7]),
                multiple ? Blank(x.Cells[8]) : null, multiple ? Blank(x.Cells[9]) : null);
        }).ToList();
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(RcpUpSaveRequest r, string user)
    {
        if (r.Rows.Count == 0) throw new BusinessRuleException(NotUploaded);
        if (!r.Multiple && string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(NoReceiptNo);
        if (!r.Multiple && string.IsNullOrWhiteSpace(r.MedixReceiptNo)) throw new BusinessRuleException(NoMedixNo);
        var date = ClinicInvoiceRules.Date(r.ReceiptDate);
        var amount = Amount(r.ReceiptAmt);
        await using var tx = await sessions.BeginAsync();
        // The match is made again on the server: only rows whose claim exists are saved, as the desktop's "Matched" rows.
        var claims = new Dictionary<string, (string CasNumber, string? BordxRefNo, decimal CLMBordxAmt)>(StringComparer.OrdinalIgnoreCase);
        foreach (var chunk in r.Rows.Select(x => x.ClaimNo.Trim()).Where(x => x.Length > 0).Distinct().Chunk(500))
            foreach (var c in await repository.ClaimsAsync(tx, string.Join(',', chunk))) claims.TryAdd(c.CasNumber, c);
        foreach (var row in r.Rows)
        {
            if (!claims.TryGetValue(row.ClaimNo.Trim(), out var c)) continue;
            var receiptNo = r.Multiple ? (row.ReceiptNo ?? "").Trim() : r.ReceiptNo!.Trim();
            var medix = r.Multiple ? Blank(row.MedixReceiptNo) : r.MedixReceiptNo!.Trim();
            await repository.SaveAsync(tx, c.CasNumber, Blank(c.BordxRefNo), receiptNo, date, amount, Blank(r.ChequeNo), row.Amount, Blank(r.Month), Blank(r.Year),
                Blank(row.ClientPvNo), medix, user);
        }
        await tx.CommitAsync();
        return Updated;
    }

    private static decimal Amount(string? v) => decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0;

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
}
