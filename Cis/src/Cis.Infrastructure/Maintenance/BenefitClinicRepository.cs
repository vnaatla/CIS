using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.BenefitClinic;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IBenefitClinicRepository"/> on the <c>cis.BenefitClinic_*</c> procedures.</summary>
public sealed class BenefitClinicRepository : IBenefitClinicRepository
{
    /// <inheritdoc />
    public Task<List<LookupOption>> CategoriesAsync(IDbSession session) => session.QueryAsync<LookupOption>("cis.BenefitClinic_Categories");

    /// <inheritdoc />
    public Task<List<ClinicBenefitFee>> SearchAsync(IDbSession session, string cardType, DateTime effDate) =>
        session.QueryAsync<ClinicBenefitFee>("cis.BenefitClinic_Search", new { cardType, effDate });

    /// <inheritdoc />
    public Task<bool> ExistsAsync(IDbSession session, string cardType, DateTime effDate) =>
        session.ScalarAsync<bool>("cis.BenefitClinic_Exists", new { cardType, effDate });

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, string cardType, DateTime effDate, string bnfCode, string? rmPerVisit, string userCode, DateTime today) =>
        session.ExecAsync("cis.BenefitClinic_Insert", new { cardType, effDate, bnfCode, rmPerVisit, user = userCode, today });
}
