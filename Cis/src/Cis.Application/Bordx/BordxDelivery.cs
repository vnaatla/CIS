using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Bordx;

/// <summary>Bordx Delivery search (FrmSendtoPayor / FrmReturnPayor / FrmSendtoAC / FrmSendtoMNRB).</summary>
public sealed record BordxDeliveryRequest
{
#pragma warning disable CS1591
    /// <summary>Group company name (payor for Send to MNRB: "code - name").</summary>
    public string? Company { get; init; }
    public bool Scan { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public string? DateFrom { get; init; }
    public string? DateTo { get; init; }
    public string? AmountFrom { get; init; }
    public string? AmountTo { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
    /// <summary>0 Not Sent, 1 Sent, 2 Both.</summary>
    public int? Status { get; init; }
    /// <summary>Send to A/C: date sent from / to.</summary>
    public string? SentFrom { get; init; }
    public string? SentTo { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class BordxDeliveryRequestValidator : Common.RequestValidator<BordxDeliveryRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxDeliveryRequestValidator()
    {
        RuleFor(r => r.Company).MaximumLength(200);
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.AmountFrom).MaximumLength(20);
        RuleFor(r => r.AmountTo).MaximumLength(20);
        RuleFor(r => r.MonthFrom).MaximumLength(4);
        RuleFor(r => r.MonthTo).MaximumLength(4);
        RuleFor(r => r.YearFrom).MaximumLength(6);
        RuleFor(r => r.YearTo).MaximumLength(6);
    }
}

/// <summary>Update: the ticked bordereaux and the date.</summary>
public sealed record BordxDeliveryUpdateRequest
{
    /// <summary>Ticked bordx numbers.</summary>
    public List<string> BordxNos { get; init; } = [];
    /// <summary>dd/mm/yyyy.</summary>
    public string? Date { get; init; }
}

/// <summary>Lengths.</summary>
public sealed class BordxDeliveryUpdateRequestValidator : Common.RequestValidator<BordxDeliveryUpdateRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxDeliveryUpdateRequestValidator() => RuleFor(r => r.BordxNos).Must(b => b.Count <= 30000 && b.All(x => x.Length <= 20));
}

/// <summary>The listed bordereaux (Done = already sent / returned) and the total amount.</summary>
public record BordxDeliveryResult(List<Dictionary<string, object?>> Rows, decimal TotalAmount);

/// <summary>cis.BordxDelivery_Search parameters.</summary>
public record BordxDeliveryFilter(string Kind, string? Company, string? PayCode, bool Scan, string? BordxFrom, string? BordxTo, DateTime? DateFrom, DateTime? DateTo,
    string? AmountFrom, string? AmountTo, string? MonthFrom, string? MonthTo, string? YearFrom, string? YearTo, string? Status, DateTime? SentFrom, DateTime? SentTo);

/// <summary>Data access.</summary>
public interface IBordxDeliveryRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, BordxDeliveryFilter filter);
    Task<int> UpdateAsync(IDbSession session, string kind, string bordxNo, DateTime date);
    Task<(List<string> Companies, List<(string Code, string Name)> Payors)> LookupsAsync(IDbSession session);
#pragma warning restore CS1591
}

/// <summary>Claim Bordx > Bordx Delivery (the four screens, by kind: payor, returned, ac, mnrb).</summary>
public interface IBordxDeliveryService
{
    /// <summary>ComboListing.</summary>
    Task<(List<string> Companies, List<(string Code, string Name)> Payors)> LookupsAsync();

    /// <summary>cmdSearch_Click.</summary>
    Task<BordxDeliveryResult> SearchAsync(string kind, BordxDeliveryRequest request);

    /// <summary>CmdToPayor / CmdFrPayor / CmdToAC / CmdToMNRB (after the confirmation).</summary>
    Task<string> UpdateAsync(string kind, BordxDeliveryUpdateRequest request);
}

/// <inheritdoc />
public sealed class BordxDeliveryService(IDbSessions sessions, IBordxDeliveryRepository repository) : IBordxDeliveryService
{
    /// <summary>The screens.</summary>
    public static readonly string[] Kinds = ["payor", "returned", "ac", "mnrb"];

    /// <summary>Desktop texts.</summary>
    public const string NoCompany = "Please select a Group Company", NoPayor = "Please select a Payor", NoBordx = "Please Enter Bordx No", NoRecord = "No record found!",
        NoneTicked = "Please select at least one record to be updated! ", NoDate = "Please Enter in date", Updated = "Record updated...",
        EarlierThanSent = "Date Received from Payor cannot be ealier than Date Sent to Payor!", Done = "Record has been Updated!";

    /// <inheritdoc />
    public async Task<(List<string> Companies, List<(string Code, string Name)> Payors)> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <inheritdoc />
    public async Task<BordxDeliveryResult> SearchAsync(string kind, BordxDeliveryRequest r)
    {
        Kind(kind);
        if (r.Scan)
        {
            if (string.IsNullOrWhiteSpace(r.Company)) throw new BusinessRuleException(kind == "mnrb" ? NoPayor : NoCompany);
            if (string.IsNullOrWhiteSpace(r.BordxFrom)) throw new BusinessRuleException(NoBordx);
        }
        var company = Blank(r.Company);
        var payCode = kind == "mnrb" && company is not null ? (company.IndexOf('-') is > 0 and var i ? company[..i].Trim() : company[..Math.Min(2, company.Length)]) : null;
        var filter = new BordxDeliveryFilter(kind, kind == "mnrb" ? null : company, payCode, r.Scan, Blank(r.BordxFrom), r.Scan ? null : Blank(r.BordxTo),
            ClinicInvoiceRules.Date(r.DateFrom), ClinicInvoiceRules.Date(r.DateTo), Blank(r.AmountFrom), Blank(r.AmountTo), Blank(r.MonthFrom), Blank(r.MonthTo),
            Blank(r.YearFrom), Blank(r.YearTo), r.Status is 0 or 1 ? r.Status.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : null,
            ClinicInvoiceRules.Date(r.SentFrom), ClinicInvoiceRules.Date(r.SentTo));
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, filter);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new BordxDeliveryResult(rows, rows.Sum(x => x.GetValueOrDefault("BordxClaimAmt") is { } v and not DBNull ? Convert.ToDecimal(v, System.Globalization.CultureInfo.InvariantCulture) : 0));
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(string kind, BordxDeliveryUpdateRequest r)
    {
        Kind(kind);
        var nos = r.BordxNos.Select(x => x.Trim()).Where(x => x.Length > 0).Distinct().ToList();
        if (nos.Count == 0) throw new BusinessRuleException(NoneTicked);
        var date = ClinicInvoiceRules.Date(r.Date) ?? throw new BusinessRuleException(NoDate);
        await using var tx = await sessions.BeginAsync();
        var result = 1;
        foreach (var no in nos) result = await repository.UpdateAsync(tx, kind, no, date);
        await tx.CommitAsync();
        // Returned Fr Grp. Company: the desktop says so when the date is earlier than the date sent (the other ticked bordereaux are still updated).
        return result == 2 ? EarlierThanSent : Updated;
    }

    private static void Kind(string kind)
    {
        if (!Kinds.Contains(kind)) throw new NotFoundException("Unknown delivery screen.");
    }

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
}
