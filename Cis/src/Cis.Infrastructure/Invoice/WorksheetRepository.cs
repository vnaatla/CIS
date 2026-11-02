using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IWorksheetRepository"/> on <c>cis.Worksheet_Lookups</c>.</summary>
public sealed class WorksheetRepository : IWorksheetRepository
{
    private sealed record Clinic(string? BenefitType, string? ReductionStatus, decimal McoProcFeePercent);
    private sealed record RateRow(string Code, decimal? Rate);
    private sealed record Category(string Code, string? Description);
    private sealed record Reduction(decimal Amount, decimal Percent);

    /// <inheritdoc />
    public async Task<WorksheetLookups> LookupsAsync(IDbSession session, WorksheetInput i)
    {
        using var g = await session.MultipleAsync("cis.Worksheet_Lookups", new
        {
            clinicCode = i.ClinicCode.Trim(), card = i.ClinicBenefitCard.Trim(), planCode = i.PlanCode.Trim(), insCode = i.InsCode.Trim(), hltCode = i.HltCode.Trim(),
        });
        var categories = (await g.ReadAsync<Category>()).Select(c => (c.Code, c.Description ?? "")).ToList();
        var clinic = (await g.ReadAsync<Clinic>()).FirstOrDefault();
        var std = (await g.ReadAsync<RateRow>()).GroupBy(r => r.Code).ToDictionary(x => x.Key, x => x.First().Rate);
        var spe = (await g.ReadAsync<RateRow>()).GroupBy(r => r.Code).ToDictionary(x => x.Key, x => x.First().Rate);
        var plan = await g.ReadSingleAsync<bool>();
        var benefits = (await g.ReadAsync<WorksheetBenefit>()).ToList();
        var red = (await g.ReadAsync<Reduction>()).FirstOrDefault();
        return new WorksheetLookups
        {
            Categories = categories, ClinicBenefitType = clinic?.BenefitType ?? "", ClinicReductionStatus = clinic?.ReductionStatus ?? "",
            McoProcFeePercent = clinic?.McoProcFeePercent ?? 0, StandardRates = std, SpecialRates = spe, PlanBenefitStatus = plan, Benefits = benefits,
            ReductionAmount = red?.Amount ?? 0, ReductionPercent = red?.Percent ?? 0,
        };
    }
}
