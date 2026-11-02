using Cis.Application.Abstractions;
using Cis.Application.Claim;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Claim;

/// <summary><see cref="IPassToClaimRepository"/> on the <c>cis.Pass2Claim_*</c> procedures.</summary>
public sealed class PassToClaimRepository : IPassToClaimRepository
{
    /// <inheritdoc />
    public async Task<List<string>> UsersAsync(IDbSession session) =>
        (await session.RowsAsync("cis.Pass2Claim_Users")).Select(r => (r.GetValueOrDefault("CASUserToken") as string)?.Trim() ?? "")
            .Where(u => u.Length > 0).ToList();

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PassFilters f, string? casNumbers)
    {
        static string? T(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? First(string? v) => T(v) is { } t ? t[..1].ToUpperInvariant() : null;
        var caseNo = T(f.CaseNo)?.ToUpperInvariant();
        return session.RowsAsync("cis.Pass2Claim_Search", new
        {
            caseNo = f.Scan ? caseNo : null, casFrom = f.Scan ? null : caseNo, casTo = f.Scan ? null : T(f.CaseNo2), caseStatus = First(f.CaseStatus),
            claimability = First(f.Claimability), passBy = T(f.PassBy), passed = First(f.Status), noVisit = f.NoVisit, noRegister = f.NoRegister, noPass = f.NoPass,
            regFrom = PassToClaimService.Date(f.RegFrom), regTo = PassToClaimService.Date(f.RegTo), visitFrom = PassToClaimService.Date(f.VisitFrom),
            visitTo = PassToClaimService.Date(f.VisitTo), passFrom = PassToClaimService.Date(f.PassFrom), passTo = PassToClaimService.Date(f.PassTo), casNumbers,
        });
    }

    /// <inheritdoc />
    public Task PassAsync(IDbSession session, string casNumber, string userCode, DateTime today) =>
        session.RowsAsync("cis.Pass2Claim_Pass", new { casNumber, userCode, today });
}
