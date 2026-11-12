using Cis.Application.Abstractions;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Reports;

/// <summary>Tracking > Grp. Com. Clinic search (FrmTracking2 SearchRecords). Mode 1 by group, 2 clinics per group, 3 by clinic, 4 groups per clinic.</summary>
public sealed record TrkGrpClnRequest(int Mode, string? GroupCode, string? ClinicCode, string? GroupCompany, string? Clinic);

/// <summary>Lengths.</summary>
public sealed class TrkGrpClnRequestValidator : Common.RequestValidator<TrkGrpClnRequest>
{
    /// <summary>Defines the rules.</summary>
    public TrkGrpClnRequestValidator()
    {
        RuleFor(r => r.Mode).InclusiveBetween(1, 4);
        RuleFor(r => r.GroupCode).MaximumLength(30);
        RuleFor(r => r.ClinicCode).MaximumLength(30);
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.Clinic).MaximumLength(200);
    }
}

/// <summary>Data access.</summary>
public interface ITrkGrpClnRepository
{
#pragma warning disable CS1591
    Task<(List<string> Companies, List<string> Clinics)> LookupsAsync(IDbSession session);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, TrkGrpClnRequest request);
#pragma warning restore CS1591
}

/// <summary>Tracking > Grp. Com. Clinic (FrmTracking2).</summary>
public interface ITrkGrpClnService
{
    /// <summary>Group company and clinic names (Check1_Click).</summary>
    Task<(List<string> Companies, List<string> Clinics)> LookupsAsync();

    /// <summary>cmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(TrkGrpClnRequest request);
}

/// <inheritdoc />
public sealed class TrkGrpClnService(IDbSessions sessions, ITrkGrpClnRepository repository) : ITrkGrpClnService
{
    /// <summary>Desktop text.</summary>
    public const string NoRecord = "No record found!";

    /// <inheritdoc />
    public async Task<(List<string> Companies, List<string> Clinics)> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(TrkGrpClnRequest r)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r with { GroupCode = B(r.GroupCode), ClinicCode = B(r.ClinicCode), GroupCompany = B(r.GroupCompany), Clinic = B(r.Clinic) });
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return rows;
    }
}
