using Cis.Application.Abstractions;

namespace Cis.Application.Lookups;

/// <summary>A drop-down entry (code + description), e.g. the desktop's "PAYCode - PAYCompanyName" combos.</summary>
/// <param name="Code">Stored code.</param>
/// <param name="Name">Description shown to the user.</param>
public record LookupOption(string Code, string Name);

/// <summary>Data access of the shared drop-down lists.</summary>
public interface ILookupRepository
{
    /// <summary>CISDB.PAY (optionally of one health type), as <c>Select PAYCode, PAYCompanyName from PAY [Where HLTCode = ...]</c>.</summary>
    Task<List<LookupOption>> PayorsAsync(IDbSession session, string? healthType);

    /// <summary>CISDB.HLT.</summary>
    Task<List<LookupOption>> HealthTypesAsync(IDbSession session);

    /// <summary>CISDB.INS.</summary>
    Task<List<LookupOption>> InsuredTypesAsync(IDbSession session);

    /// <summary>CISDB.GRPCompany (optionally of one payor), as <c>select GRPCoName from GRPCompany where GRPCoPAYCode = ...</c>.</summary>
    Task<List<LookupOption>> GroupCompaniesAsync(IDbSession session, string? payor);

    /// <summary>HISDB.CLN; <paramref name="panelOnly"/> = <c>CLNPanel IN ('Y','C')</c>; <paramref name="search"/> matches code or name.</summary>
    Task<List<LookupOption>> ClinicsAsync(IDbSession session, bool panelOnly, string? search, int max);
}

/// <summary>Shared drop-down lists used by many screens.</summary>
public interface ILookupService
{
    /// <summary>Payors, optionally of one health type.</summary>
    Task<List<LookupOption>> PayorsAsync(string? healthType);

    /// <summary>Health types.</summary>
    Task<List<LookupOption>> HealthTypesAsync();

    /// <summary>Insured types.</summary>
    Task<List<LookupOption>> InsuredTypesAsync();

    /// <summary>Group companies, optionally of one payor.</summary>
    Task<List<LookupOption>> GroupCompaniesAsync(string? payor);

    /// <summary>Clinics (at most <see cref="LookupService.MaxClinics"/>).</summary>
    Task<List<LookupOption>> ClinicsAsync(bool panelOnly, string? search);
}

/// <inheritdoc />
public sealed class LookupService(IDbSessions sessions, ILookupRepository repository) : ILookupService
{
    /// <summary>Most clinics returned by one search (HISDB.CLN has 8,000+ rows).</summary>
    public const int MaxClinics = 500;

    /// <inheritdoc />
    public async Task<List<LookupOption>> PayorsAsync(string? healthType)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.PayorsAsync(s, healthType);
    }

    /// <inheritdoc />
    public async Task<List<LookupOption>> HealthTypesAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.HealthTypesAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<LookupOption>> InsuredTypesAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.InsuredTypesAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<LookupOption>> GroupCompaniesAsync(string? payor)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.GroupCompaniesAsync(s, payor);
    }

    /// <inheritdoc />
    public async Task<List<LookupOption>> ClinicsAsync(bool panelOnly, string? search)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.ClinicsAsync(s, panelOnly, search, MaxClinics);
    }
}
