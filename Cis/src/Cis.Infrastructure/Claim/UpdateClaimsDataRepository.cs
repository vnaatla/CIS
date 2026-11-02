using Cis.Application.Abstractions;
using Cis.Application.Claim;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Claim;

/// <summary><see cref="IUpdateClaimsDataRepository"/> on the <c>cis.ClaimData_*</c> procedures.</summary>
public sealed class UpdateClaimsDataRepository : IUpdateClaimsDataRepository
{
    /// <inheritdoc />
    public Task<int> UpdateAsync(IDbSession session, ClaimDataKind kind, string key, decimal? amount, DateTime? date, DateTime today, string userCode) =>
        kind == ClaimDataKind.BordxDate
            ? session.ScalarAsync<int>("cis.ClaimData_BordxDate", new { bordxRefNo = key, date, today, userCode })
            : session.ScalarAsync<int>($"cis.ClaimData_{kind}", new { casNumber = key, amount, today, userCode });
}
