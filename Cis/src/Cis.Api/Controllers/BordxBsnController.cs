using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Bordx;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > Print Bordx > BSN-PrintSummary (frmBSNPrintSummary).</summary>
[Route("api/claim-bordx/bsn-summary")]
[RequireAccess(AccessRight.BordxPrint)]
public sealed class BordxBsnController(IBordxBsnService bsn) : ApiControllerBase
{
    /// <summary>The BSN bordereaux summary for the bordx numbers.</summary>
    [HttpGet]
    public Task<List<Dictionary<string, object?>>> Summary([FromQuery] string[]? bordxNo) => bsn.SummaryAsync(bordxNo ?? []);
}
