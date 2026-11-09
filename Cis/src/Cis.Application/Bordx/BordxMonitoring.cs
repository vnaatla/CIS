using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Bordx;

/// <summary>Bordx Control filters (FrmBordxControl).</summary>
public sealed record BordxControlRequest
{
#pragma warning disable CS1591
    public string? Company { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public string? DateFrom { get; init; }
    public string? DateTo { get; init; }
    /// <summary>Sent to payor: 1 sent, 0 not sent (check boxes), null either.</summary>
    public string? Sent { get; init; }
    /// <summary>Docs returned: 1 / 0.</summary>
    public string? Returned { get; init; }
    /// <summary>Paid: F full, P partial, U unpaid.</summary>
    public string? Paid { get; init; }
    public string? Subsidiary { get; init; }
    /// <summary>C Cashless, M Member.</summary>
    public string? CashType { get; init; }
    /// <summary>0 Health Connect, 1 Sihat Johor.</summary>
    public int Entity { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class BordxControlRequestValidator : Common.RequestValidator<BordxControlRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxControlRequestValidator()
    {
        RuleFor(r => r.Company).MaximumLength(200);
        RuleFor(r => r.Subsidiary).MaximumLength(200);
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
    }
}

/// <summary>Bordx &amp; Worksheet filters (FrmBordxWks).</summary>
public sealed record BordxWksRequest
{
#pragma warning disable CS1591
    public string? Company { get; init; }
    /// <summary>"code - name".</summary>
    public string? Payor { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    public string? ReceiptFrom { get; init; }
    public string? ReceiptTo { get; init; }
    public string? RecFrom { get; init; }
    public string? RecTo { get; init; }
    public bool NoRecDate { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class BordxWksRequestValidator : Common.RequestValidator<BordxWksRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxWksRequestValidator()
    {
        RuleFor(r => r.Company).MaximumLength(200);
        RuleFor(r => r.Payor).MaximumLength(200);
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
        RuleFor(r => r.ReceiptFrom).MaximumLength(30);
        RuleFor(r => r.ReceiptTo).MaximumLength(30);
    }
}

/// <summary>A list and its total (Bordx Control: the MCO fee column, as the desktop; Outstanding: total due; Bordx &amp; Worksheet: approved).</summary>
public record BordxListResult(List<Dictionary<string, object?>> Rows, decimal Total, decimal Total14 = 0, decimal TotalOver14 = 0);

/// <summary>Data access.</summary>
public interface IBordxMonitoringRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> ControlAsync(IDbSession session, BordxControlRequest r, DateTime? from, DateTime? to, string? payTo);
    Task<List<Dictionary<string, object?>>> OutstandingAsync(IDbSession session, string? company);
    Task<List<Dictionary<string, object?>>> WorksheetsAsync(IDbSession session, BordxWksRequest r, string? payCode, DateTime? recFrom, DateTime? recTo);
    Task<(List<string> Companies, List<(string Code, string Name)> Payors, List<string> Subsidiaries)> LookupsAsync(IDbSession session, string? company);
#pragma warning restore CS1591
}

/// <summary>Claim Bordx > Monitoring: Bordx Control, Outstanding Bordx, Bordx &amp; Worksheet.</summary>
public interface IBordxMonitoringService
{
    /// <summary>Drop-downs (subsidiaries of a group company).</summary>
    Task<(List<string> Companies, List<(string Code, string Name)> Payors, List<string> Subsidiaries)> LookupsAsync(string? company);

    /// <summary>FrmBordxControl cmdSearch_Click.</summary>
    Task<BordxListResult> ControlAsync(BordxControlRequest request);

    /// <summary>FrmOSBordx cmdSearch_Click: days outstanding up to the O/S date, split at 14 days.</summary>
    Task<BordxListResult> OutstandingAsync(string? company, string? osDate);

    /// <summary>FrmBordxWks cmdSearch_Click.</summary>
    Task<BordxListResult> WorksheetsAsync(BordxWksRequest request);
}

/// <inheritdoc />
public sealed class BordxMonitoringService(IDbSessions sessions, IBordxMonitoringRepository repository, TimeProvider clock) : IBordxMonitoringService
{
    /// <summary>Desktop texts.</summary>
    public const string NoRecord = "No record found!", NoRecordCompany = "No record found for this Group Company!";

    /// <inheritdoc />
    public async Task<(List<string> Companies, List<(string Code, string Name)> Payors, List<string> Subsidiaries)> LookupsAsync(string? company)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s, Blank(company));
    }

    /// <inheritdoc />
    public async Task<BordxListResult> ControlAsync(BordxControlRequest r)
    {
        await using var s = await sessions.OpenAsync();
        var payTo = (r.CashType ?? "").Trim() is { Length: > 0 } c ? c[..1].ToUpperInvariant() : null;
        var rows = await repository.ControlAsync(s, r with { Company = Blank(r.Company), Subsidiary = Blank(r.Subsidiary), BordxFrom = Blank(r.BordxFrom), BordxTo = Blank(r.BordxTo) },
            ClinicInvoiceRules.Date(r.DateFrom), ClinicInvoiceRules.Date(r.DateTo), payTo is "C" or "M" ? payTo : null);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        foreach (var x in rows)
        {
            foreach (var k in new[] { "NoofDay", "NoofDay2" })
                if (x.GetValueOrDefault(k) is { } v and not DBNull && Convert.ToInt32(v, CultureInfo.InvariantCulture) == 0) x[k] = 1;
        }
        return new BordxListResult(rows, rows.Sum(x => N(x, "CLMMCOProcFee")));
    }

    /// <inheritdoc />
    public async Task<BordxListResult> OutstandingAsync(string? company, string? osDate)
    {
        // The O/S date starts as today on the desktop; an invalid date leaves the days blank.
        var os = string.IsNullOrWhiteSpace(osDate) ? clock.GetLocalNow().DateTime.Date : ClinicInvoiceRules.Date(osDate);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.OutstandingAsync(s, Blank(company));
        if (rows.Count == 0) throw new NotFoundException(NoRecordCompany);
        decimal within = 0, over = 0;
        var n = 0;
        foreach (var x in rows)
        {
            x["No"] = ++n;
            int? days = os is { } o && x.GetValueOrDefault("BordxPAYSentDate") is DateTime sent ? (int)(o.Date - sent.Date).TotalDays + 1 : null;
            if (days == 0) days = 1;
            x["Days"] = days;
            var amt = N(x, "BordxClaimAmt");
            // The desktop puts the amount in the "> 14 days" column when there are no days (blank compares as greater).
            if (days is { } d && d <= 14) { x["Within14"] = amt; within += amt; }
            else { x["Over14"] = amt; over += amt; }
        }
        return new BordxListResult(rows, within + over, within, over);
    }

    /// <inheritdoc />
    public async Task<BordxListResult> WorksheetsAsync(BordxWksRequest r)
    {
        var payor = Blank(r.Payor);
        var payCode = payor is null ? null : payor.IndexOf('-') is > 0 and var i ? payor[..i].Trim() : payor[..Math.Min(4, payor.Length)];
        await using var s = await sessions.OpenAsync();
        var rows = (await repository.WorksheetsAsync(s, r with
        {
            Company = Blank(r.Company), BordxFrom = Blank(r.BordxFrom), BordxTo = Blank(r.BordxTo), CaseFrom = Blank(r.CaseFrom), CaseTo = Blank(r.CaseTo),
            ReceiptFrom = Blank(r.ReceiptFrom), ReceiptTo = Blank(r.ReceiptTo),
        }, payCode, ClinicInvoiceRules.Date(r.RecFrom), ClinicInvoiceRules.Date(r.RecTo)))
            // The desktop lists only rows with a payor code.
            .Where(x => Convert.ToString(x.GetValueOrDefault("BordxPAYCode"), CultureInfo.InvariantCulture)?.Trim().Length > 0).ToList();
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new BordxListResult(rows, rows.Sum(x => N(x, "CLMTotalApproved")));
    }

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
