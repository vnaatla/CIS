using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Membership;

/// <summary>A staff category (CISDB WebCategory) and the plan / insured type it registers.</summary>
/// <param name="Category">Category.</param>
/// <param name="InsType">Insured type.</param>
/// <param name="PlanCode">Plan.</param>
public record StaffCategory(string Category, string InsType, string PlanCode);

/// <summary>ComboListing.</summary>
/// <param name="Races">SearchRace (code, name).</param>
/// <param name="Categories">WebCategory.</param>
/// <param name="GroupCompanies">GRPCompany names.</param>
public record RegistrationSetup(List<LookupOption> Races, List<StaffCategory> Categories, List<string> GroupCompanies);

/// <summary>A member found by IC (MBMDetails).</summary>
/// <param name="MbmNumber">Member number (shown only while the policy is active).</param>
/// <param name="Active">Policy expiry after today: Register is disabled and the member number is shown.</param>
/// <param name="Payor">Payor.</param>
/// <param name="Name">Name.</param>
/// <param name="Ic">IC.</param>
/// <param name="Race">Race code.</param>
/// <param name="Gender">Sex.</param>
/// <param name="Dob">Date of birth.</param>
/// <param name="EffDate">Policy eff date (active only).</param>
/// <param name="ExpDate">Policy exp date (active only).</param>
/// <param name="StaffCategory">"CREW" (active only).</param>
/// <param name="Department">Department.</param>
/// <param name="GroupCompany">Group company.</param>
/// <param name="EmployeeNo">Employee no.</param>
public record RegisteredMember(string? MbmNumber, bool Active, string Payor, string Name, string Ic, string Race, string Gender, DateTime? Dob,
    DateTime? EffDate, DateTime? ExpDate, string? StaffCategory, string Department, string GroupCompany, string EmployeeNo);

/// <summary>Limits of the staff category's plan (CallCLNLimit).</summary>
/// <param name="AnnualLimit">Annual limit.</param>
/// <param name="VisitLimit">Visit limit (0 when unlimited).</param>
/// <param name="VisitFull">GetVisitLimit returned "FULL" (unlimited for the DC plan).</param>
public record RegistrationLimits(decimal AnnualLimit, decimal VisitLimit, bool VisitFull);

/// <summary>CLNMemberCallAge.</summary>
/// <param name="Age">GetAge ("34Y", "-1D", …).</param>
/// <param name="AgeBand">"code - description".</param>
/// <param name="AdultChild">C for band 01, otherwise A.</param>
public record RegistrationAge(string Age, string AgeBand, string AdultChild);

/// <summary>Registration of a clinic member.</summary>
/// <param name="Ic">Patient IC.</param>
/// <param name="Name">Member Name.</param>
/// <param name="Race">Race Code.</param>
/// <param name="Gender">Gender (M / F).</param>
/// <param name="StaffCategory">Staff Category.</param>
/// <param name="EmployeeNo">Employee No.</param>
/// <param name="EffDate">Effective Date dd/mm/yyyy.</param>
/// <param name="ExpDate">Expiry Date dd/mm/yyyy.</param>
/// <param name="Dob">Birth Date dd/mm/yyyy.</param>
/// <param name="GroupCompany">Group Company.</param>
/// <param name="Department">Department.</param>
public record RegistrationRequest(string? Ic, string? Name, string? Race, string? Gender, string? StaffCategory, string? EmployeeNo, string? EffDate,
    string? ExpDate, string? Dob, string? GroupCompany, string? Department);

/// <summary>Column lengths.</summary>
public sealed class RegistrationRequestValidator : RequestValidator<RegistrationRequest>
{
    /// <summary>Defines the rules.</summary>
    public RegistrationRequestValidator()
    {
        RuleFor(r => r.Ic).MaximumLength(30);
        RuleFor(r => r.Name).MaximumLength(200);
        RuleFor(r => r.EmployeeNo).MaximumLength(50);
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.Department).MaximumLength(200);
    }
}

/// <summary>The row cmdReg_Click writes.</summary>
/// <param name="Ic">IC.</param>
/// <param name="Name">Name.</param>
/// <param name="Sex">Sex.</param>
/// <param name="AdultChild">A / C.</param>
/// <param name="AgeCode">Age code.</param>
/// <param name="Dob">Date of birth.</param>
/// <param name="Race">Race code.</param>
/// <param name="GroupCompany">Group company.</param>
/// <param name="EmployeeNo">Employee no.</param>
/// <param name="Department">Department.</param>
/// <param name="PlanCode">Plan.</param>
/// <param name="InsType">Insured type.</param>
/// <param name="EffDate">Policy eff date.</param>
/// <param name="ExpDate">Policy exp date.</param>
/// <param name="VisitLimit">Visit limit.</param>
/// <param name="AnnualLimit">Annual limit.</param>
public record NewMember(string Ic, string Name, string Sex, string AdultChild, string AgeCode, DateTime Dob, string Race, string GroupCompany,
    string EmployeeNo, string Department, string PlanCode, string InsType, DateTime EffDate, DateTime ExpDate, decimal VisitLimit, decimal AnnualLimit);

/// <summary>Data access of the registration pop-up (<c>cis.MbmReg_*</c>, HISDB SP_SearchCISMBM_IC).</summary>
public interface IMemberRegistrationRepository
{
    /// <summary>Lists.</summary>
    Task<RegistrationSetup> SetupAsync(IDbSession session);

    /// <summary>SP_SearchCISMBM_IC first row.</summary>
    Task<Dictionary<string, object?>?> FindByIcAsync(IDbSession session, string ic);

    /// <summary>G001 departments containing the text.</summary>
    Task<List<string>> DepartmentsAsync(IDbSession session, string text);

    /// <summary>CallCLNLimit.</summary>
    Task<RegistrationLimits> LimitsAsync(IDbSession session, string plan, string insType, DateTime effDate);

    /// <summary>GetAgeBand: (code, description) or null.</summary>
    Task<(string Code, string Description)?> AgeBandAsync(IDbSession session, int age, bool days);

    /// <summary>Generates the number and writes the member; returns the member number.</summary>
    Task<string> RegisterAsync(IDbSession session, NewMember member, string userCode, DateTime today);
}

/// <summary>Clinic member registration (frmMemberRegistrationCLN).</summary>
public interface IMemberRegistrationService
{
    /// <summary>ComboListing.</summary>
    Task<RegistrationSetup> SetupAsync();

    /// <summary>cmbSearchMBM_Click.</summary>
    Task<RegisteredMember> FindByIcAsync(string? ic);

    /// <summary>CmdSearch_Click (departments).</summary>
    Task<List<string>> DepartmentsAsync(string? text);

    /// <summary>CallCLNLimit for a staff category.</summary>
    Task<RegistrationLimits> LimitsAsync(string? staffCategory, string? effDate);

    /// <summary>CLNMemberCallAge.</summary>
    Task<RegistrationAge> AgeAsync(string? effDate, string? dob);

    /// <summary>cmdReg_Click: message and member number.</summary>
    Task<(string Message, string MbmNumber)> RegisterAsync(RegistrationRequest request, string userCode);
}

/// <inheritdoc />
public sealed class MemberRegistrationService(IDbSessions sessions, IMemberRegistrationRepository repository, TimeProvider clock) : IMemberRegistrationService
{
    /// <summary>cmdReg_Click checks in order (spacing kept).</summary>
    public static readonly IReadOnlyList<(Func<RegistrationRequest, bool> Missing, string Message)> Checks =
    [
        (r => E(r.Ic), "Please enter IC No! "),
        (r => E(r.Name), "Please enter Member Name! "),
        (r => E(r.Race), "Please enter Race Code! "),
        (r => E(r.Gender), "Please enter Gender! "),
        (r => E(r.StaffCategory), "Please enter Staff Category! "),
        (r => E(r.EmployeeNo), "Please enter Employee No! "),
        (r => Date(r.EffDate) is null, "Please enter Effective Date! "),
        (r => Date(r.ExpDate) is null, "Please enter Expiry Date! "),
        (r => Date(r.Dob) is null, "Please enter Birth Date! "),
        (r => E(r.GroupCompany), "Please enter Group Company! "),
        (r => E(r.Department), "Please enter Department! "),
    ];

    /// <summary>Desktop texts.</summary>
    public const string NotFound = "No record found!", DeptFirst = "Please enter Department! ", BadDate = "Invalid Date Format! ",
        NoBand = "Age Band not found or Date out of Acceptable Range! \nPlease verify your maintenance table!", Generated = "Membership Generated, Your Membership Number is : ";

    /// <inheritdoc />
    public async Task<RegistrationSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<RegisteredMember> FindByIcAsync(string? ic)
    {
        await using var s = await sessions.OpenAsync();
        var r = await repository.FindByIcAsync(s, (ic ?? "").Trim()) ?? throw new NotFoundException(NotFound);
        var exp = r.GetValueOrDefault("MBMPolExpDate") as DateTime?;
        var active = exp > clock.GetLocalNow().Date;
        return new RegisteredMember(active ? T(r, "MBMNumber") : null, active, T(r, "PAYCode"), T(r, "MBMName"), T(r, "MBMIcBcPp"), T(r, "RACCode"),
            T(r, "MBMSex"), r.GetValueOrDefault("MBMDOB") as DateTime?, active ? r.GetValueOrDefault("MBMPolEffDate") as DateTime? : null, active ? exp : null,
            active ? "CREW" : null, T(r, "MBMDepartment"), T(r, "MBMGroupCompany"), T(r, "MBMEmployeeNo"));
    }

    /// <inheritdoc />
    public async Task<List<string>> DepartmentsAsync(string? text)
    {
        if (E(text)) throw new BusinessRuleException(DeptFirst);
        await using var s = await sessions.OpenAsync();
        return await repository.DepartmentsAsync(s, text!.Trim());
    }

    /// <inheritdoc />
    public async Task<RegistrationLimits> LimitsAsync(string? staffCategory, string? effDate)
    {
        var eff = Date(effDate) ?? throw new BusinessRuleException(BadDate);
        await using var s = await sessions.OpenAsync();
        var cat = (await repository.SetupAsync(s)).Categories.FirstOrDefault(c => c.Category == (staffCategory ?? "").Trim());
        return cat is null || cat.PlanCode.Length == 0 ? new RegistrationLimits(0, 0, false) : await repository.LimitsAsync(s, cat.PlanCode, cat.InsType, eff);
    }

    /// <inheritdoc />
    public async Task<RegistrationAge> AgeAsync(string? effDate, string? dob)
    {
        var eff = Date(effDate) ?? throw new BusinessRuleException(BadDate);
        var birth = Date(dob) ?? throw new BusinessRuleException(BadDate);
        await using var s = await sessions.OpenAsync();
        return await Age(s, eff, birth);
    }

    /// <inheritdoc />
    public async Task<(string Message, string MbmNumber)> RegisterAsync(RegistrationRequest r, string userCode)
    {
        foreach (var (missing, message) in Checks)
            if (missing(r)) throw new BusinessRuleException(message);
        var (eff, exp, dob) = (Date(r.EffDate)!.Value, Date(r.ExpDate)!.Value, Date(r.Dob)!.Value);
        await using var s = await sessions.BeginAsync();
        var cat = (await repository.SetupAsync(s)).Categories.FirstOrDefault(c => c.Category == r.StaffCategory!.Trim());
        var (plan, ins) = (cat?.PlanCode ?? "", cat?.InsType ?? "");
        var limits = plan.Length == 0 ? new RegistrationLimits(0, 0, false) : await repository.LimitsAsync(s, plan, ins, eff);
        var age = await Age(s, eff, dob);
        var member = new NewMember(r.Ic!.Trim(), r.Name!.Trim(), r.Gender!.Trim(), age.AdultChild, age.AgeBand.Length >= 2 ? age.AgeBand[..2] : age.AgeBand,
            dob, r.Race!.Trim(), r.GroupCompany!.Trim(), r.EmployeeNo!.Trim(), r.Department!.Trim(), plan, ins, eff, exp, limits.VisitLimit, limits.AnnualLimit);
        var number = await repository.RegisterAsync(s, member, userCode, clock.GetLocalNow().Date);
        await s.CommitAsync();
        return (Generated + number, number);
    }

    private async Task<RegistrationAge> Age(IDbSession s, DateTime eff, DateTime dob)
    {
        var age = GetAge(eff, dob);
        var days = age.EndsWith('D');
        var band = await repository.AgeBandAsync(s, int.Parse(age[..^1], CultureInfo.InvariantCulture), days) ?? throw new BusinessRuleException(NoBand);
        return new RegistrationAge(age, $"{band.Code} - {band.Description}", band.Code == "01" ? "C" : "A");
    }

    /// <summary>
    /// Medix.bas GetAge as written: up to 17 calendar years the age is "-1D" (same year or later) or "0Y" / "-1Y" (the year count is never
    /// set in that branch); otherwise completed years, one less when the birth month is later or, failing that, the birth day is later.
    /// </summary>
    public static string GetAge(DateTime effectiveDate, DateTime dateOfBirth)
    {
        var checkDate = effectiveDate.Year - dateOfBirth.Year;
        if (checkDate <= 0) return "-1D";
        var years = checkDate <= 17 ? 0 : checkDate;
        var younger = dateOfBirth.Month > effectiveDate.Month || dateOfBirth.Day > effectiveDate.Day;
        return $"{(younger ? years - 1 : years).ToString(CultureInfo.InvariantCulture)}Y";
    }

    private static DateTime? Date(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static bool E(string? v) => string.IsNullOrWhiteSpace(v);

    private static string T(Dictionary<string, object?> row, string key) =>
        row.TryGetValue(key, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";
}
