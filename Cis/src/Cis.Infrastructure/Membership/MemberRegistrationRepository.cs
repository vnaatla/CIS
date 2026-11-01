using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Membership;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Membership;

/// <summary><see cref="IMemberRegistrationRepository"/> on the <c>cis.MbmReg_*</c> procedures and HISDB SP_SearchCISMBM_IC.</summary>
public sealed class MemberRegistrationRepository : IMemberRegistrationRepository
{
    /// <inheritdoc />
    public async Task<RegistrationSetup> SetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.MbmReg_Setup");
        var races = (await grid.ReadAsync<(string RACCode, string RACName)>()).Select(r => new LookupOption(r.RACCode.Trim(), r.RACName?.Trim() ?? "")).ToList();
        var categories = (await grid.ReadAsync<StaffCategory>()).ToList();
        var groups = (await grid.ReadAsync<string>()).ToList();
        return new RegistrationSetup(races, categories, groups);
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> FindByIcAsync(IDbSession session, string ic) =>
        (await session.RowsAsync("cis.Hdb_SP_SearchCISMBM_IC", new { ICNumber = ic })).FirstOrDefault();

    /// <inheritdoc />
    public Task<List<string>> DepartmentsAsync(IDbSession session, string text) => session.QueryAsync<string>("cis.MbmReg_Departments", new { text });

    /// <inheritdoc />
    public async Task<RegistrationLimits> LimitsAsync(IDbSession session, string plan, string insType, DateTime effDate) =>
        (await session.QueryAsync<RegistrationLimits>("cis.MbmReg_Limits", new { plnCode = plan, insCode = insType, effDate })).Single();

    /// <inheritdoc />
    public async Task<(string Code, string Description)?> AgeBandAsync(IDbSession session, int age, bool days)
    {
        var r = (await session.RowsAsync("cis.MbmReg_AgeBand", new { age, days })).FirstOrDefault();
        return r is null ? null : ((string)r["AgeCode"]!, r["AgeDescription"] as string ?? "");
    }

    /// <inheritdoc />
    public async Task<string> RegisterAsync(IDbSession session, NewMember m, string userCode, DateTime today) =>
        (await session.ScalarAsync<string>("cis.MbmReg_Register", new
        {
            ic = m.Ic, name = m.Name, sex = m.Sex, adultChild = m.AdultChild, ageCode = m.AgeCode, dob = m.Dob, race = m.Race, groupCompany = m.GroupCompany,
            employeeNo = m.EmployeeNo, department = m.Department, plnCode = m.PlanCode, insCode = m.InsType, effDate = m.EffDate, expDate = m.ExpDate,
            visitLimit = m.VisitLimit, annualLimit = m.AnnualLimit, user = userCode, today,
        }))!;
}
