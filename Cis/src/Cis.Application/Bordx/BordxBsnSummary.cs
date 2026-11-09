using Cis.Application.Abstractions;
using Cis.Domain.Errors;

namespace Cis.Application.Bordx;

/// <summary>Data access.</summary>
public interface IBordxBsnRepository
{
    /// <summary>cis.BordxBsn_Summary.</summary>
    Task<List<Dictionary<string, object?>>> SummaryAsync(IDbSession session, string bordxNos);
}

/// <summary>Claim Bordx > Print Bordx > BSN-PrintSummary (frmBSNPrintSummary).</summary>
public interface IBordxBsnService
{
    /// <summary>cmdsearch_Click for the entered bordx numbers.</summary>
    Task<List<Dictionary<string, object?>>> SummaryAsync(IReadOnlyList<string> bordxNos);
}

/// <inheritdoc />
public sealed class BordxBsnService(IDbSessions sessions, IBordxBsnRepository repository) : IBordxBsnService
{
    /// <summary>Desktop text.</summary>
    public const string NoBordx = "Please Enter Bordx Refno..! ";

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SummaryAsync(IReadOnlyList<string> bordxNos)
    {
        var list = bordxNos.Select(b => b.Trim()).Where(b => b.Length > 0).Distinct().Take(200).ToList();
        if (list.Count == 0 || list.Any(b => b.Length > 20 || b.Contains(','))) throw new BusinessRuleException(NoBordx);
        await using var s = await sessions.OpenAsync();
        return await repository.SummaryAsync(s, string.Join(',', list));
    }
}
