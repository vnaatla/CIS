using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Lookups;

/// <inheritdoc />
public sealed class LookupRepository : ILookupRepository
{
    /// <inheritdoc />
    public Task<List<LookupOption>> PayorsAsync(IDbSession session, string? healthType) =>
        session.QueryAsync<LookupOption>("cis.Lookup_Payors", new { healthType = Blank(healthType) });

    /// <inheritdoc />
    public Task<List<LookupOption>> HealthTypesAsync(IDbSession session) => session.QueryAsync<LookupOption>("cis.Lookup_HealthTypes");

    /// <inheritdoc />
    public Task<List<LookupOption>> InsuredTypesAsync(IDbSession session) => session.QueryAsync<LookupOption>("cis.Lookup_InsuredTypes");

    /// <inheritdoc />
    public Task<List<LookupOption>> GroupCompaniesAsync(IDbSession session, string? payor) =>
        session.QueryAsync<LookupOption>("cis.Lookup_GroupCompanies", new { payor = Blank(payor) });

    /// <inheritdoc />
    public Task<List<LookupOption>> ClinicsAsync(IDbSession session, bool panelOnly, string? search, int max) =>
        session.QueryAsync<LookupOption>("cis.Lookup_Clinics",
            new { panelOnly, search = Blank(search) is { } s ? "%" + SqlText.EscapeLike(s) + "%" : null, max });

    private static string? Blank(string? text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim();
}
