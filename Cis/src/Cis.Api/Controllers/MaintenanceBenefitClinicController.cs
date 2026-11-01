using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.BenefitClinic;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Maintenance > Benefit > Clinic (frmMaintenanceBenefitsCln; frmMaintenanceBenefitsClnNew uses the same calls).</summary>
[Route("api/maintenance/benefit-clinic")]
[RequireAccess(AccessRight.MaintenanceBenefitClinic)]
public sealed class MaintenanceBenefitClinicController(IBenefitClinicService clinic) : ApiControllerBase
{
    /// <summary>Card types and the grid rows (categories).</summary>
    [HttpGet("setup")]
    public async Task<object> Setup() => new { cardTypes = clinic.CardTypes, categories = await clinic.CategoriesAsync() };

    /// <summary>Saved RM Payable Fees of a card type + effective date (dd/mm/yyyy).</summary>
    [HttpGet]
    public Task<List<ClinicBenefitFee>> Search([FromQuery] string? cardType, [FromQuery] string? effDate) => clinic.SearchAsync(cardType, effDate);

    /// <summary>Add + Save.</summary>
    [HttpPost]
    public async Task<object> Save(ClinicBenefitSaveRequest request) => new { message = await clinic.SaveAsync(request, UserCode) };
}
