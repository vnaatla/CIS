using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Claim;
using Cis.Domain.Errors;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > Update Claims Data (frmUpdateClaimsData): spreadsheet verify / upload.</summary>
[Route("api/claim-bordx/update-claims-data")]
[RequireAccess(AccessRight.BordxPrint)]
public sealed class UpdateClaimsDataController(IUpdateClaimsDataService data) : ApiControllerBase
{
    /// <summary>Verify the file (kind: Ffs, Copay, BordxAmt, BordxDate).</summary>
    [HttpPost("{kind}/verify")]
    [RequestSizeLimit(20_000_000)]
    public ClaimDataResult Verify(ClaimDataKind kind, IFormFile? file)
    {
        using var stream = Open(file);
        return data.Verify(kind, stream);
    }

    /// <summary>Upload the file.</summary>
    [HttpPost("{kind}/upload")]
    [RequestSizeLimit(20_000_000)]
    public async Task<ClaimDataResult> Upload(ClaimDataKind kind, IFormFile? file)
    {
        await using var stream = Open(file);
        return await data.UploadAsync(kind, stream, UserCode);
    }

    private static Stream Open(IFormFile? file)
    {
        if (file is null || file.Length == 0) throw new BusinessRuleException(UpdateClaimsDataService.FileFirst);
        var copy = new MemoryStream();
        file.CopyTo(copy);
        copy.Position = 0;
        return copy;
    }
}
