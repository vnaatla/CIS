using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Bordx;

/// <summary>UpdateBordxDate.</summary>
public sealed record BordxDateUpdateRequest
{
    /// <summary>The bordx no.</summary>
    public string BordxNo { get; init; } = "";
    /// <summary>The new bordx date (dd/mm/yyyy).</summary>
    public string? BordxDate { get; init; }
}

/// <summary>Lengths.</summary>
public sealed class BordxDateUpdateRequestValidator : Common.RequestValidator<BordxDateUpdateRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxDateUpdateRequestValidator() => RuleFor(r => r.BordxNo).MaximumLength(10);
}

/// <summary>One line of BordxDateChange (per posting status).</summary>
public record BordxDateRow(string BordxRefNo, DateTime? BordxDate, string PostingStatus, int NoOfCases);

/// <summary>Data access.</summary>
public interface IBordxDateRepository
{
#pragma warning disable CS1591
    Task<List<BordxDateRow>> FindAsync(IDbSession session, string bordxNo);
    Task UpdateAsync(IDbSession session, string bordxNo, DateTime date, string user);
#pragma warning restore CS1591
}

/// <summary>Claim Bordx > ClaimBordxDate (FrmBordxDate).</summary>
public interface IBordxDateService
{
    /// <summary>CmdSearch_Click: the lines and whether posting is done (the desktop then says so).</summary>
    Task<(List<BordxDateRow> Rows, bool Posted)> FindAsync(string? bordxNo);

    /// <summary>CmdUpdate_Click.</summary>
    Task<string> UpdateAsync(BordxDateUpdateRequest request, string user);
}

/// <inheritdoc />
public sealed class BordxDateService(IDbSessions sessions, IBordxDateRepository repository) : IBordxDateService
{
    /// <summary>Desktop texts.</summary>
    public const string NotFound = "Record not found", NoDate = "Please Enter Bordx Date", Posted = "Posting is Done. Update is not allowed",
        Updated = "BordxDate Updated Successfully..!!";

    /// <inheritdoc />
    public async Task<(List<BordxDateRow> Rows, bool Posted)> FindAsync(string? bordxNo)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.FindAsync(s, (bordxNo ?? "").Trim());
        if (rows.Count == 0) throw new NotFoundException(NotFound);
        return (rows, rows.Any(r => r.PostingStatus == "YES"));
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(BordxDateUpdateRequest r, string user)
    {
        var date = ClinicInvoiceRules.Date(r.BordxDate) ?? throw new BusinessRuleException(NoDate);
        await using var tx = await sessions.BeginAsync();
        // The desktop uses the posting status of the last search; with no lines nothing is posted, so the update runs (and changes nothing).
        if ((await repository.FindAsync(tx, r.BordxNo.Trim())).Any(x => x.PostingStatus == "YES")) throw new BusinessRuleException(Posted);
        await repository.UpdateAsync(tx, r.BordxNo.Trim(), date, user);
        await tx.CommitAsync();
        return Updated;
    }
}
