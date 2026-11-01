using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Membership;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Clinic member registration pop-up (frmMemberRegistrationCLN), opened from the clinic invoice.</summary>
[Route("api/membership/registration")]
[RequireAccess(AccessRight.InvoiceClinic, AccessRight.InvoiceMember)]
public sealed class MemberRegistrationController(IMemberRegistrationService registration) : ApiControllerBase
{
    /// <summary>Races, staff categories, group companies.</summary>
    [HttpGet("setup")]
    public Task<RegistrationSetup> Setup() => registration.SetupAsync();

    /// <summary>Search MBM by the patient IC.</summary>
    [HttpGet("member")]
    public Task<RegisteredMember> Member([FromQuery] string? ic) => registration.FindByIcAsync(ic);

    /// <summary>Departments of G001 containing the text.</summary>
    [HttpGet("departments")]
    public Task<List<string>> Departments([FromQuery] string? text) => registration.DepartmentsAsync(text);

    /// <summary>Annual / visit limit of the staff category's plan.</summary>
    [HttpGet("limits")]
    public Task<RegistrationLimits> Limits([FromQuery] string? staffCategory, [FromQuery] string? effDate) => registration.LimitsAsync(staffCategory, effDate);

    /// <summary>Age, age band and adult / child.</summary>
    [HttpGet("age")]
    public Task<RegistrationAge> Age([FromQuery] string? effDate, [FromQuery] string? dob) => registration.AgeAsync(effDate, dob);

    /// <summary>Register.</summary>
    [HttpPost]
    public async Task<object> Register(RegistrationRequest request)
    {
        var (message, number) = await registration.RegisterAsync(request, UserCode);
        return new { message, mbmNumber = number };
    }
}
