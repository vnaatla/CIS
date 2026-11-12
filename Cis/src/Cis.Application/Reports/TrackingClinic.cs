using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Reports;

/// <summary>Tracking > Clinic search (FrmTracking cmdSearch_Click). The combos carry "code - text".</summary>
public sealed record TrkClnRequest
{
#pragma warning disable CS1591
    public string? Code { get; init; }
    public string? Name { get; init; }
    public string? Group { get; init; }
    public string? Postcode { get; init; }
    public string? City { get; init; }
    public string? State { get; init; }
    public string? Contact { get; init; }
    public string? Doctor { get; init; }
    public string? Initial { get; init; }
    public string? BnfType { get; init; }
    public string? ReceivedFrom { get; init; }
    public string? ReceivedTo { get; init; }
    public string? Panel { get; init; }
    public string? Type { get; init; }
    public string? Xray { get; init; }
    public string? ForRef { get; init; }
    public string? Location { get; init; }
    public string? Ffs { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class TrkClnRequestValidator : Common.RequestValidator<TrkClnRequest>
{
    /// <summary>Defines the rules.</summary>
    public TrkClnRequestValidator()
    {
        RuleFor(r => r.Code).MaximumLength(10);
        RuleFor(r => r.Name).MaximumLength(200);
        RuleFor(r => r.Group).MaximumLength(300);
        RuleFor(r => r.City).MaximumLength(100);
        RuleFor(r => r.State).MaximumLength(100);
        RuleFor(r => r.Contact).MaximumLength(200);
        RuleFor(r => r.Doctor).MaximumLength(200);
        RuleFor(r => r.Initial).MaximumLength(100);
        RuleFor(r => r.ForRef).MaximumLength(50);
        RuleFor(r => r.Location).MaximumLength(200);
    }
}

/// <summary>The parsed criteria.</summary>
public sealed record TrkClnFilter(string? Code, string? Name, string? Group, string? Postcode, string? City, string? State, string? Contact, string? Doctor,
    string? Initial, string? BnfType, DateTime? ReceivedFrom, DateTime? ReceivedTo, string? Panel, string? Type, string? Xray, string? ForRef, string? Location,
    string? Ffs);

/// <summary>Data access.</summary>
public interface ITrkClnRepository
{
#pragma warning disable CS1591
    Task<(List<string> Groups, List<string> Cities, List<string> States)> LookupsAsync(IDbSession session);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, TrkClnFilter filter);
#pragma warning restore CS1591
}

/// <summary>Tracking > Clinic (FrmTracking).</summary>
public interface ITrkClnService
{
    /// <summary>Clinic groups, cities and states.</summary>
    Task<(List<string> Groups, List<string> Cities, List<string> States)> LookupsAsync();

    /// <summary>cmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(TrkClnRequest request);
}

/// <inheritdoc />
public sealed class TrkClnService(IDbSessions sessions, ITrkClnRepository repository) : ITrkClnService
{
    /// <summary>Desktop texts.</summary>
    public const string NoRecord = "Record not found", BadDate = "Invalid Date Format! ";

    /// <inheritdoc />
    public async Task<(List<string> Groups, List<string> Cities, List<string> States)> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(TrkClnRequest r)
    {
        var group = Blank(r.Group);
        if (group is not null) group = group.Contains('-') ? group[..group.IndexOf('-')].Trim() : group[..Math.Min(4, group.Length)];
        var filter = new TrkClnFilter(Blank(r.Code), Blank(r.Name), group, Blank(r.Postcode), Blank(r.City), Blank(r.State), Blank(r.Contact), Blank(r.Doctor),
            Blank(r.Initial), First(r.BnfType), Date(r.ReceivedFrom), Date(r.ReceivedTo), First(r.Panel), First(r.Type), First(r.Xray), Blank(r.ForRef), Blank(r.Location),
            First(r.Ffs));
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, filter);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return rows;
    }

    private static DateTime? Date(string? v) => string.IsNullOrWhiteSpace(v) ? null : ClinicInvoiceRules.Date(v) ?? throw new BusinessRuleException(BadDate);

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string? First(string? v) => Blank(v)?[..1];
}
