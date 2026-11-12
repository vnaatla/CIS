using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.AcExport;

/// <summary>Invoice Posting search (frmInvPosting SearchRecord).</summary>
public sealed record InvPostRequest(string? Payor, string? InvoiceFrom, string? InvoiceTo, string? RegisterFrom, string? RegisterTo, string? Posted);

/// <summary>Lengths.</summary>
public sealed class InvPostRequestValidator : Common.RequestValidator<InvPostRequest>
{
    /// <summary>Defines the rules.</summary>
    public InvPostRequestValidator()
    {
        RuleFor(r => r.Payor).MaximumLength(300);
        RuleFor(r => r.InvoiceFrom).MaximumLength(100);
        RuleFor(r => r.InvoiceTo).MaximumLength(100);
        RuleFor(r => r.Posted).MaximumLength(20);
    }
}

/// <summary>Data access.</summary>
public interface IInvPostRepository
{
#pragma warning disable CS1591
    Task<(List<Dictionary<string, object?>> Summary, List<Dictionary<string, object?>> Detail)> SearchAsync(IDbSession session, string? payCode, string? invFrom,
        string? invTo, DateTime? regFrom, DateTime? regTo, string? posted);
    Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, string? payCode, string? invFrom, string? invTo, DateTime? regFrom, DateTime? regTo, string posted);
    Task MarkAsync(IDbSession session, string invoiceNo, string clinicCode, string user);
#pragma warning restore CS1591
}

/// <summary>A/C Export > Invoice Posting (frmInvPosting).</summary>
public interface IInvPostService
{
    /// <summary>cmdSearch_Click.</summary>
    Task<PostingResult> SearchAsync(InvPostRequest request);

    /// <summary>cmdPost_Click: the posting workbook; the listed cases not yet posted are marked posted.</summary>
    Task<FileOutput> PostAsync(InvPostRequest request, string user);
}

/// <inheritdoc />
public sealed class InvPostService(IDbSessions sessions, IInvPostRepository repository, IPostInvClnService export, TimeProvider clock) : IInvPostService
{
    /// <summary>Desktop texts.</summary>
    public const string BadDate = "Invalid date format!", FutureDate = "Invoice Date cannot greater than system date ! ";

    /// <inheritdoc />
    public async Task<PostingResult> SearchAsync(InvPostRequest r)
    {
        var (payCode, from, to, posted) = Criteria(r);
        await using var s = await sessions.OpenAsync();
        var (summary, detail) = await repository.SearchAsync(s, payCode, Blank(r.InvoiceFrom), Blank(r.InvoiceTo), from, to, posted);
        return Lines(summary, detail);
    }

    /// <inheritdoc />
    public async Task<FileOutput> PostAsync(InvPostRequest r, string user)
    {
        var (payCode, from, to, posted) = Criteria(r);
        await using var tx = await sessions.BeginAsync();
        var rows = await repository.RowsAsync(tx, payCode, Blank(r.InvoiceFrom), Blank(r.InvoiceTo), from, to, posted ?? "0");
        var lines = new List<PostingLine>();
        foreach (var x in rows)
        {
            lines.Add(Line(x, "6000", "D", T(x, "BordxRefNo")));
            lines.Add(Line(x, "84" + T(x, "ClinicCode"), "C", T(x, "CLNName")));
            // UpdateCase only marks cases when the "not posted" status is searched.
            if ((posted ?? "0") == "0") await repository.MarkAsync(tx, T(x, "InvoiceNo"), T(x, "ClinicCode"), user);
        }
        await tx.CommitAsync();
        var total = rows.Sum(x => N(x, "CLMPayableByMedix2C"));
        return export.Export(new PostingResult(lines, total, total, rows.Count, false), "InvoicePosting");
    }

    /// <summary>INVPostListing: per summary row a credit line (84 + clinic) and a debit line (60 + group company code for payor DC, else 60 + payor) per matching
    /// detail row.</summary>
    public static PostingResult Lines(List<Dictionary<string, object?>> summary, List<Dictionary<string, object?>> detail)
    {
        var lines = new List<PostingLine>();
        decimal debit = 0, credit = 0;
        var byKey = detail.ToLookup(d => (T(d, "SumInvoiceNo"), d.GetValueOrDefault("SumRegisterDate") as DateTime?));
        foreach (var x in summary)
        {
            lines.Add(Line(x, "84" + T(x, "ClinicCode"), "C", T(x, "InvoiceNo")));
            credit += N(x, "CLMPayableByMedix2C");
            foreach (var d in byKey[(T(x, "InvoiceNo"), x.GetValueOrDefault("RegisterDate") as DateTime?)])
            {
                lines.Add(Line(d, "60" + (T(d, "PAYCode") == "DC" ? T(d, "GRPCoCode") : T(d, "PAYCode")), "D", T(d, "BordxRefNo")));
                debit += N(d, "CLMPayableByMedix2C");
            }
        }
        return new PostingResult(lines, debit, credit, summary.Count, false);
    }

    private static PostingLine Line(IReadOnlyDictionary<string, object?> x, string account, string dc, string description)
    {
        var inv = x.GetValueOrDefault("InvoiceDate") as DateTime?;
        var reg = x.GetValueOrDefault("RegisterDate") as DateTime?;
        var bordx = T(x, "BordxRefNo");
        var posted = T(x, "CASInvPost");
        return new PostingLine(posted == "1", "RC", "", "", inv?.ToString("MM/yyyy", CultureInfo.InvariantCulture) ?? "",
            inv?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "", T(x, "InvoiceNo"), account, N(x, "CLMPayableByMedix2C"), dc, "", T(x, "PAYCode"),
            T(x, "ClinicCode"), false, inv, description)
        {
            Extra = new Dictionary<string, string>
            {
                ["RegPeriod"] = reg?.ToString("yyMM", CultureInfo.InvariantCulture) ?? "", ["GrpCo"] = T(x, "GRPCoCode"),
                ["Bordx"] = bordx.Length > 3 ? bordx.Substring(3, Math.Min(5, bordx.Length - 3)) : "",
                ["PostedText"] = posted == "0" ? "No" : posted.Length > 0 ? "Yes" : "", ["PostUser"] = T(x, "CASInvPostUser"),
                ["PostDate"] = (x.GetValueOrDefault("CASInvPostDate") as DateTime?)?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "",
            },
        };
    }

    private (string? PayCode, DateTime? From, DateTime? To, string? Posted) Criteria(InvPostRequest r)
    {
        var payor = (r.Payor ?? "").Trim();
        var posted = (r.Posted ?? "").Trim();
        return (payor.Length > 0 ? payor[..Math.Min(2, payor.Length)] : null, Date(r.RegisterFrom), Date(r.RegisterTo), posted.Length > 0 ? posted[..1] : null);
    }

    private DateTime? Date(string? v)
    {
        if (string.IsNullOrWhiteSpace(v)) return null;
        var d = ClinicInvoiceRules.Date(v) ?? throw new BusinessRuleException(BadDate);
        if (d.Date > clock.GetLocalNow().Date) throw new BusinessRuleException(FutureDate);
        return d;
    }

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
