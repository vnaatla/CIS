using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Member info of the invoice screens (MemberSuppInfo, FrmINVMBMinfo).</summary>
[Route("api/invoice/member-info")]
[RequireAccess(AccessRight.InvoiceClinic, AccessRight.InvoiceMember, AccessRight.InvoiceAdjustment, AccessRight.InvoiceAdjustmentAfterBordx)]
public sealed class MemberSuppController(IMemberSuppService members) : ApiControllerBase
{
    /// <summary>The member's plan and limit details.</summary>
    [HttpGet("{mbmNumber}/{covId}")]
    public Task<MemberSupp> Info(string mbmNumber, string covId) => members.InfoAsync(mbmNumber, covId);
}
