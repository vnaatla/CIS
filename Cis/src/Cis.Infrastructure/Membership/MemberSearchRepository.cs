using Cis.Application.Abstractions;
using Cis.Application.Membership;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Membership;

/// <summary><see cref="IMemberSearchRepository"/> on the <c>cis.MemberSearch_*</c> procedures.</summary>
public sealed class MemberSearchRepository : IMemberSearchRepository
{
    /// <inheritdoc />
    public Task<List<MemberSearchRow>> PrincipalsAsync(IDbSession session, MemberSearchBy by, string text, int max) =>
        Rows(session, "cis.MemberSearch_Principals", new { by = by.ToString(), text, max });

    /// <inheritdoc />
    public Task<List<MemberSearchRow>> SuppsAsync(IDbSession session, MemberSearchBy by, string text, int max) =>
        Rows(session, "cis.MemberSearch_Supps", new { by = by.ToString(), text, max });

    /// <inheritdoc />
    public Task<List<MemberSearchRow>> ByPrefixAsync(IDbSession session, string prefix) =>
        Rows(session, "cis.MemberSearch_ByPrefix", new { prefix });

    private static async Task<List<MemberSearchRow>> Rows(IDbSession session, string procedure, object args) =>
        (await session.RowsAsync(procedure, args)).Select(r => new MemberSearchRow(
            Text(r, "MBMNumber") ?? "", Text(r, "CoverId"), Text(r, "Name"), Text(r, "Ic"), Text(r, "PolicyNo"), Text(r, "PolicyStatus"),
            Text(r, "DataStatus"), r.GetValueOrDefault("EffDate") as DateTime?, r.GetValueOrDefault("ExpDate") as DateTime?, Text(r, "GroupCompany"),
            Text(r, "Department"), r.GetValueOrDefault("Supp") is true)).ToList();

    private static string? Text(Dictionary<string, object?> row, string key) =>
        row.TryGetValue(key, out var v) && v is not null ? Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture)?.Trim() : null;
}
