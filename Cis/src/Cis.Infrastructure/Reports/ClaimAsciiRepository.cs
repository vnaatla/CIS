using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="IClaimAsciiRepository"/>: cis.ClaimAscii_* and the CISWorksheet ClaimDetails_ASCII* procedures.</summary>
public sealed class ClaimAsciiRepository : IClaimAsciiRepository
{
    /// <inheritdoc />
    public Task<List<string>> ClinicsAsync(IDbSession session, bool bintulu) => session.QueryAsync<string>("cis.ClaimAscii_Clinics", new { bintulu });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, AsciiFormat format, string? bordxFrom, string? bordxTo, string? clinic) => format switch
    {
        AsciiFormat.Bintulu => session.RowsAsync("cis.ClaimAscii_Bintulu", new { bordxFrom, bordxTo, clinic }, 300),
        AsciiFormat.Toshiba or AsciiFormat.Singer =>
            session.RowsAsync("CISWorksheet.dbo.ClaimDetails_ASCII", new { BordxRefNoFrom = bordxFrom, BordxRefNoTo = bordxTo }, 300),
        AsciiFormat.Tims => session.RowsAsync("CISWorksheet.dbo.ClaimDetails_ASCII_TIMS", new { BordxRefNoFrom = bordxFrom }, 300),
        AsciiFormat.Cimb => session.RowsAsync("CISWorksheet.dbo.ClaimDetails_ASCII_CIMB", new { BordxRefNoFrom = bordxFrom }, 300),
        AsciiFormat.OmniProviders => session.RowsAsync("cis.ClaimAscii_OmniPanel", null, 300),
        _ => session.RowsAsync("CISWorksheet.dbo.ClaimDetails_ASCII_OMNI", new { BordxRefNoFrom = bordxFrom }, 300),
    };

    /// <inheritdoc />
    public async Task<string?> DeductibleAsync(IDbSession session, string plan, string marital) =>
        (await session.RowsAsync("cis.ClaimAscii_Deductible", new { plan, marital })).Select(r => r["Amount"]).FirstOrDefault() is { } v and not DBNull
            ? Convert.ToDecimal(v, System.Globalization.CultureInfo.InvariantCulture).ToString("0.############", System.Globalization.CultureInfo.InvariantCulture)
            : null;

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> OmniProvidersAsync(IDbSession session) => session.RowsAsync("cis.ClaimAscii_OmniProviders");

    /// <inheritdoc />
    public Task OmniSentAsync(IDbSession session, string clinic) => session.ExecAsync("cis.ClaimAscii_OmniSent", new { clinic });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> UtusanAsync(IDbSession session, int claimType, string bordxFrom, string bordxTo) =>
        session.RowsAsync("cis.ClaimAscii_Utusan", new { @case = claimType, bordxFrom, bordxTo }, 300);
}
