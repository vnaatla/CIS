using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Payment;

/// <summary>Pay Clinic Post search (frmPayClinicPost SearchRecord).</summary>
public sealed record PayClnPostRequest(string? Payor, string? InvoiceFrom, string? InvoiceTo, string? RegisterFrom, string? RegisterTo, string? Posted);

/// <summary>Lengths.</summary>
public sealed class PayClnPostRequestValidator : Common.RequestValidator<PayClnPostRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayClnPostRequestValidator()
    {
        RuleFor(r => r.Payor).MaximumLength(300);
        RuleFor(r => r.InvoiceFrom).MaximumLength(50);
        RuleFor(r => r.InvoiceTo).MaximumLength(50);
        RuleFor(r => r.Posted).MaximumLength(20);
    }
}

/// <summary>A posting line (lstCLNPosting).</summary>
public sealed record PayClnPostLine(string Account, string Period, decimal Amount, string DebitCredit, string Date, string Journal, string Reference, string Description,
    string ClinicCode, string Blank, string PayCode, string RegisterPeriod, string GrpCoCode, string Bordx, string Posted, string PostUser, string PostDate);

/// <summary>The lines and the debit / credit totals.</summary>
public sealed record PayClnPostResult(List<PayClnPostLine> Lines, decimal Debit, decimal Credit);

/// <summary>Data access.</summary>
public interface IPayClnPostRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? payCode, string? invFrom, string? invTo, DateTime? regFrom, DateTime? regTo,
        string? posted);
    Task MarkAsync(IDbSession session, string invoiceNo, string clinicCode, string user);
#pragma warning restore CS1591
}

/// <summary>Payment > Pay Clinic Post (frmPayClinicPost).</summary>
public interface IPayClnPostService
{
    /// <summary>cmdSearch_Click.</summary>
    Task<PayClnPostResult> SearchAsync(PayClnPostRequest request);

    /// <summary>cmdPost_Click: the posting workbook; the cases not yet posted are marked posted.</summary>
    Task<FileOutput> PostAsync(PayClnPostRequest request, string user);
}

/// <inheritdoc />
public sealed class PayClnPostService(IDbSessions sessions, IPayClnPostRepository repository, IReportExporter exporter, TimeProvider clock) : IPayClnPostService
{
    /// <summary>Desktop texts.</summary>
    public const string BadDate = "Invalid date format!", FutureDate = "Invoice Date cannot greater than system date ! ", NoRecord = "No record found!";

    /// <inheritdoc />
    public async Task<PayClnPostResult> SearchAsync(PayClnPostRequest r)
    {
        await using var s = await sessions.OpenAsync();
        return Lines(await Rows(s, r));
    }

    private async Task<List<Dictionary<string, object?>>> Rows(IDbSession s, PayClnPostRequest r)
    {
        var from = Date(r.RegisterFrom);
        var to = Date(r.RegisterTo);
        var payor = (r.Payor ?? "").Trim();
        var posted = (r.Posted ?? "").Trim();
        return await repository.SearchAsync(s, payor.Length > 0 ? payor[..Math.Min(2, payor.Length)] : null, Blank(r.InvoiceFrom), Blank(r.InvoiceTo), from, to,
            posted.Length > 0 ? posted[..1] : null);
    }

    private DateTime? Date(string? v)
    {
        if (string.IsNullOrWhiteSpace(v)) return null;
        var d = ClinicInvoiceRules.Date(v) ?? throw new BusinessRuleException(BadDate);
        if (d.Date > clock.GetLocalNow().Date) throw new BusinessRuleException(FutureDate);
        return d;
    }

    /// <summary>INVPostListing: a debit line (84 + clinic) and a credit line (7600) per row.</summary>
    public static PayClnPostResult Lines(List<Dictionary<string, object?>> rows)
    {
        var lines = new List<PayClnPostLine>();
        decimal debit = 0, credit = 0;
        foreach (var x in rows)
        {
            var pvDate = x.GetValueOrDefault("PVCLNDate") as DateTime?;
            var reg = x.GetValueOrDefault("RegisterDate") as DateTime?;
            var amount = N(x, "CLNPaidAmt");
            var bordx = T(x, "BordxRefNo");
            var posted = T(x, "CASPayCLNPost") == "0" ? "No" : T(x, "CASPayCLNPost").Length > 0 ? "Yes" : "";
            var postDate = (x.GetValueOrDefault("CASPayCLNPostDate") as DateTime?)?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "";
            string P(string f) => pvDate?.ToString(f, CultureInfo.InvariantCulture) ?? "";
            var common = (Reg: reg?.ToString("yyMM", CultureInfo.InvariantCulture) ?? "", Bordx: bordx.Length > 3 ? bordx.Substring(3, Math.Min(5, bordx.Length - 3)) : "");
            lines.Add(new PayClnPostLine("84" + T(x, "ClinicCode"), P("MM/yyyy"), amount, "D", P("dd/MM/yyyy"), "CP", T(x, "PVCLNNo"), T(x, "InvoiceNo"), T(x, "ClinicCode"), "",
                T(x, "PAYCode"), common.Reg, T(x, "GRPCoCode"), common.Bordx, posted, T(x, "CASPayCLNPostUser"), postDate));
            lines.Add(new PayClnPostLine("7600", P("MM/yyyy"), amount, "C", P("dd/MM/yyyy"), "CP", T(x, "InvoiceNo"), $"{T(x, "PVCLNChequeNo")} - {T(x, "CLNName")}",
                T(x, "ClinicCode"), "", T(x, "PAYCode"), common.Reg, T(x, "GRPCoCode"), common.Bordx, posted, "", postDate));
            debit += amount;
            credit += amount;
        }
        return new PayClnPostResult(lines, debit, credit);
    }

    /// <inheritdoc />
    public async Task<FileOutput> PostAsync(PayClnPostRequest r, string user)
    {
        await using var tx = await sessions.BeginAsync();
        var rows = await Rows(tx, r);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var result = Lines(rows);
        foreach (var (inv, cln) in rows.Where(x => T(x, "CASPayCLNPost") == "0" && T(x, "InvoiceNo").Length > 0)
                     .Select(x => (T(x, "InvoiceNo"), T(x, "ClinicCode"))).Distinct())
            await repository.MarkAsync(tx, inv, cln, user);
        await tx.CommitAsync();
        var data = result.Lines.Select(l => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
        {
            ["A"] = l.Account, ["B"] = l.Period, ["C"] = l.Amount, ["D"] = l.DebitCredit, ["E"] = l.Date, ["F"] = l.Journal, ["G"] = l.Reference, ["H"] = l.Description,
            ["I"] = l.ClinicCode, ["J"] = l.Blank, ["K"] = l.PayCode, ["L"] = l.RegisterPeriod, ["M"] = l.GrpCoCode, ["N"] = l.Bordx,
        }).ToList();
        return exporter.Export(new ReportDocument("PAY CLINIC POSTING",
        [
            new("A", "Account"), new("B", "Period"), new("C", "Amount", ReportFormat.Money, true), new("D", "D/C"), new("E", "Date"), new("F", "Journal"),
            new("G", "Reference"), new("H", "Description"), new("I", "Clinic"), new("J", ""), new("K", "Payor"), new("L", "Reg Period"), new("M", "Grp Co"),
            new("N", "Bordx"),
        ], data) { FileName = "PayClinicPosting" }, ReportFileType.Excel);
    }

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
