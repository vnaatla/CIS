using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Bordx;

/// <summary>Remove Worksheet: the ticked worksheets of a bordereaux.</summary>
public sealed record BordxRemoveRequest
{
    /// <summary>The bordx no. on the screen.</summary>
    public string BordxNo { get; init; } = "";
    /// <summary>The ticked case numbers.</summary>
    public List<string> Cases { get; init; } = [];
}

/// <summary>Lengths.</summary>
public sealed class BordxRemoveRequestValidator : Common.RequestValidator<BordxRemoveRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxRemoveRequestValidator()
    {
        RuleFor(r => r.BordxNo).MaximumLength(20);
        RuleFor(r => r.Cases).Must(c => c.Count <= 30000 && c.All(x => x.Length <= 30));
    }
}

/// <summary>The bordereaux's worksheets with its dates and totals.</summary>
public record BordxRemoveResult(List<Dictionary<string, object?>> Rows, DateTime? PrintedDate, DateTime? BordxDate, int Total, decimal TotalApproved);

/// <summary>Data access.</summary>
public interface IBordxRemoveRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> ListAsync(IDbSession session, string bordxNo);
    Task RemoveAsync(IDbSession session, string bordxNo, string casNumber, string user);
#pragma warning restore CS1591
}

/// <summary>Claim Bordx > Worksheet > Remove Worksheet (FrmBordxRevise).</summary>
public interface IBordxRemoveService
{
    /// <summary>CmdGo_Click.</summary>
    Task<BordxRemoveResult> FindAsync(string? bordxNo);

    /// <summary>CmdReject_Click (after the confirmation).</summary>
    Task<string> RemoveAsync(BordxRemoveRequest request, string user);
}

/// <inheritdoc />
public sealed class BordxRemoveService(IDbSessions sessions, IBordxRemoveRepository repository) : IBordxRemoveService
{
    /// <summary>Desktop texts.</summary>
    public const string NoBordx = "Please select a Bordx No ", NoRecord = "No record found!", Updated = "Record updated...",
        Confirm = "You are going to remove a worksheet from current Bordereaux. \nAre you sure?", NoneTicked = "Please select a worksheet to remove.";

    /// <inheritdoc />
    public async Task<BordxRemoveResult> FindAsync(string? bordxNo)
    {
        if (string.IsNullOrWhiteSpace(bordxNo)) throw new BusinessRuleException(NoBordx);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.ListAsync(s, bordxNo.Trim());
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var total = rows.Sum(x => x.GetValueOrDefault("CLMTotalApproved") is { } v and not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0);
        return new BordxRemoveResult(rows, rows[0].GetValueOrDefault("BordxPrintedDate") as DateTime?, rows[0].GetValueOrDefault("BordxDate") as DateTime?, rows.Count, total);
    }

    /// <inheritdoc />
    public async Task<string> RemoveAsync(BordxRemoveRequest r, string user)
    {
        if (string.IsNullOrWhiteSpace(r.BordxNo)) throw new BusinessRuleException(NoBordx);
        var cases = r.Cases.Select(c => c.Trim()).Where(c => c.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (cases.Count == 0) throw new BusinessRuleException(NoneTicked);
        await using var tx = await sessions.BeginAsync();
        var listed = (await repository.ListAsync(tx, r.BordxNo.Trim())).Select(x => Convert.ToString(x["CASNumber"], CultureInfo.InvariantCulture)!.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var c in cases.Where(listed.Contains)) await repository.RemoveAsync(tx, r.BordxNo.Trim(), c, user);
        await tx.CommitAsync();
        return Updated;
    }
}
