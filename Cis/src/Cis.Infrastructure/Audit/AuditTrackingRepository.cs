using Cis.Application.Abstractions;
using Cis.Application.Audit;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Audit;

/// <summary><see cref="IAuditTrackingRepository"/> on cis.AuditTrack_Search / cis.AuditTrack_Lookups.</summary>
public sealed class AuditTrackingRepository : IAuditTrackingRepository
{
    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, AuditTrackingRequest f, AuditTrackingDates d, decimal? amountFrom, decimal? amountTo) =>
        await session.RowsAsync("cis.AuditTrack_Search", new
        {
            caseFrom = V(f.CaseFrom), caseTo = V(f.CaseTo), caseBlank = f.CaseBlank, payTo = V(f.PayTo), payToBlank = f.PayToBlank,
            clinic = V(f.Clinic), clinicBlank = f.ClinicBlank, payor = V(f.Payor), payorBlank = f.PayorBlank,
            groupCompany = V(f.GroupCompany), groupBlank = f.GroupBlank, patientName = V(f.PatientName), nameBlank = f.NameBlank,
            mbmFrom = V(f.MemberFrom), mbmTo = V(f.MemberTo), mbmBlank = f.MemberBlank, invFrom = V(f.InvoiceFrom), invTo = V(f.InvoiceTo), invBlank = f.InvoiceBlank,
            bordxFrom = V(f.BordxFrom), bordxTo = V(f.BordxTo), bordxBlank = f.BordxBlank, caseStatus = V(f.CaseStatus), claimability = V(f.Claimability),
            planFrom = V(f.PlanFrom), planTo = V(f.PlanTo), planBlank = f.PlanBlank, bordxDateFrom = d.BordxFrom, bordxDateTo = d.BordxTo,
            icFrom = V(f.IcFrom), icTo = V(f.IcTo), icBlank = f.IcBlank, visitFrom = d.VisitFrom, visitTo = d.VisitTo, visitBlank = f.VisitBlank,
            regFrom = d.RegisterFrom, regTo = d.RegisterTo, regBlank = f.RegisterBlank,
            invDateFrom = d.InvoiceFrom, invDateTo = d.InvoiceTo, invDateBlank = f.InvoiceDateBlank,
            enteredFrom = d.EnteredFrom, enteredTo = d.EnteredTo, enteredBlank = f.EnteredBlank, amtFrom = amountFrom, amtTo = amountTo, amtBlank = f.AmountBlank,
            monthFrom = V(f.MonthFrom), monthTo = V(f.MonthTo), monthBlank = f.MonthBlank, yearFrom = V(f.YearFrom), yearTo = V(f.YearTo), yearBlank = f.YearBlank,
            clinicType = V(f.ClinicType), opFrom = V(f.OpBatchFrom), opTo = V(f.OpBatchTo), opBlank = f.OpBatchBlank, ffs = V(f.Ffs), omni = f.Omni, tims = f.Tims,
            subsidiary = V(f.Subsidiary), department = V(f.Department), omniSent = V(f.OmniSent), audited = f.Audited,
        }, timeout: 300);

    /// <inheritdoc />
    public async Task<(List<string> Subsidiaries, List<string> Departments)> LookupsAsync(IDbSession session, string groupCompany)
    {
        using var grid = await session.MultipleAsync("cis.AuditTrack_Lookups", new { groupCompany });
        return ((await grid.ReadAsync<string>()).ToList(), (await grid.ReadAsync<string>()).ToList());
    }

    private static string? V(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
}
