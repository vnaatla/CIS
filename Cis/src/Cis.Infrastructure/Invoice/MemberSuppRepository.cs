using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IMemberSuppRepository"/> on <c>cis.MemberSupp_Info</c>.</summary>
public sealed class MemberSuppRepository : IMemberSuppRepository
{
    /// <inheritdoc />
    public async Task<MemberSuppRows> RowsAsync(IDbSession session, string mbmNumber, string covId)
    {
        using var g = await session.MultipleAsync("cis.MemberSupp_Info", new { mbmNumber, covId });
        async Task<List<Dictionary<string, object?>>> Read() =>
            (await g.ReadAsync()).Select(r => ((IDictionary<string, object?>)r).ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase)).ToList();
        var member = (await Read()).FirstOrDefault();
        var group = (await Read()).FirstOrDefault();
        var clinical = (await Read()).FirstOrDefault();
        var plan = (await Read()).FirstOrDefault();
        var copay = (await Read()).FirstOrDefault()?.GetValueOrDefault("CoAmount");
        var supp = (await Read()).FirstOrDefault()?.GetValueOrDefault("SuppLimitStatus") as string;
        var cover = (await Read()).FirstOrDefault();
        var limits = await Read();
        return new MemberSuppRows(member, group, clinical, plan, copay, supp, cover, limits);
    }
}
