using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Payment;

/// <summary>Payment Tracking search (FrmPayment / frmPayment2 SearchRecord).</summary>
public sealed record PayTrackingRequest
{
#pragma warning disable CS1591
    /// <summary>0 Pay to Clinic, 1 Pay to Member (null: not chosen).</summary>
    public int? ReportType { get; init; }
    /// <summary>By Claim, By Bordx, By PV No, By Clinic, By Group Company, General Listing.</summary>
    public string? Selection { get; init; }
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    public string? Company { get; init; }
    public string? MbmNo { get; init; }
    public string? InvoiceFrom { get; init; }
    public string? InvoiceTo { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public string? OpFrom { get; init; }
    public string? OpTo { get; init; }
    /// <summary>"code - name" (the first 4 characters are used, as the desktop).</summary>
    public string? Clinic { get; init; }
    public string? InvDateFrom { get; init; }
    public string? InvDateTo { get; init; }
    /// <summary>0 All, 1 Paid, 2 Partial Paid.</summary>
    public int Payment { get; init; }
    public string? PvFrom { get; init; }
    public string? PvTo { get; init; }
    public string? PvDateFrom { get; init; }
    public string? PvDateTo { get; init; }
    public string? AmountFrom { get; init; }
    public string? AmountTo { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayTrackingRequestValidator : Common.RequestValidator<PayTrackingRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayTrackingRequestValidator()
    {
        RuleFor(r => r.Company).MaximumLength(200);
        RuleFor(r => r.Clinic).MaximumLength(300);
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
        RuleFor(r => r.Selection).MaximumLength(30);
    }
}

/// <summary>The listing and the total paid.</summary>
public record PayTrackingResult(List<Dictionary<string, object?>> Rows, decimal TotalPaid);

/// <summary>Search parameters for cis.PayTracking_Search.</summary>
public record PayTrackingFilter(bool Member, string By, PayTrackingRequest Request, string? Clinic, DateTime?[] Dates, decimal? AmountFrom, decimal? AmountTo);

/// <summary>Data access.</summary>
public interface IPayTrackingRepository
{
    /// <summary>The listing.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayTrackingFilter filter);
}

/// <summary>Payment > Tracking / Tracking (Jupiter).</summary>
public interface IPayTrackingService
{
    /// <summary>CmdSearch_Click.</summary>
    Task<PayTrackingResult> SearchAsync(PayTrackingRequest request);
}

/// <inheritdoc />
public sealed class PayTrackingService(IDbSessions sessions, IPayTrackingRepository repository) : IPayTrackingService
{
    /// <summary>Desktop texts.</summary>
    public const string NoType = "Please select a report type.", NoRecord = "No record found!";

    private static readonly Dictionary<string, string> By = new(StringComparer.OrdinalIgnoreCase)
    {
        ["By Claim"] = "claim", ["By Bordx"] = "bordx", ["By PV No"] = "pv", ["By Clinic"] = "clinic", ["By Group Company"] = "group", ["General Listing"] = "general",
    };

    /// <inheritdoc />
    public async Task<PayTrackingResult> SearchAsync(PayTrackingRequest r)
    {
        if (r.ReportType is not (0 or 1)) throw new BusinessRuleException(NoType);
        var by = By.GetValueOrDefault((r.Selection ?? "").Trim()) ?? throw new NotFoundException(NoRecord);
        var clinic = (r.Clinic ?? "").Trim();
        var filter = new PayTrackingFilter(r.ReportType == 1, by, r, clinic.Length > 0 ? clinic[..Math.Min(4, clinic.Length)] : null,
            [ClinicInvoiceRules.Date(r.InvDateFrom), ClinicInvoiceRules.Date(r.InvDateTo), ClinicInvoiceRules.Date(r.PvDateFrom), ClinicInvoiceRules.Date(r.PvDateTo)],
            Amount(r.AmountFrom), Amount(r.AmountTo));
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, filter);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var paid = r.ReportType == 1 ? "MBMPaidAmt" : "CLNPaidAmt";
        foreach (var x in rows)
        {
            var status = Convert.ToString(x.GetValueOrDefault(r.ReportType == 1 ? "PVPayMBMStatus" : "PVPayCLNStatus"), CultureInfo.InvariantCulture)?.Trim();
            if (by == "general") x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
            if (by == "bordx" && x.GetValueOrDefault("BordxRefNo") is null or DBNull) x["BordxRefNo"] = "Empty";
        }
        return new PayTrackingResult(rows, rows.Sum(x => x.GetValueOrDefault(paid) is { } v and not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0));
    }

    private static decimal? Amount(string? v) =>
        decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;
}
