using System.Globalization;
using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim ASCII exports: {format} = bintulu, toshiba, singer, tims, cimb, omni.</summary>
[Route("api/reports/claim-ascii")]
[RequireAccess(AccessRight.ClaimAscii)]
public sealed class ClaimAsciiController(IClaimAsciiService ascii) : ApiControllerBase
{
    /// <summary>The clinic combo.</summary>
    [HttpGet("{format}/clinics")]
    public Task<List<string>> Clinics(string format) => ascii.ClinicsAsync(Parse(format));

    /// <summary>Generate: the file; the visit count and total are in the X-Total-Visits / X-Total-Amount headers.</summary>
    [HttpPost("{format}")]
    public async Task<FileContentResult> Generate(string format, ClaimAsciiRequest request)
    {
        var r = await ascii.GenerateAsync(Parse(format), request);
        Response.Headers["X-Total-Visits"] = r.Visits.ToString(CultureInfo.InvariantCulture);
        Response.Headers["X-Total-Amount"] = r.Total.ToString("0.00", CultureInfo.InvariantCulture);
        Response.Headers.AccessControlExposeHeaders = "X-Total-Visits, X-Total-Amount";
        return File(r.File.Content, r.File.ContentType, r.File.FileName);
    }

    private static AsciiFormat Parse(string format) =>
        Enum.TryParse<AsciiFormat>(format, true, out var f) ? f : throw new NotFoundException("Unknown export.");
}
