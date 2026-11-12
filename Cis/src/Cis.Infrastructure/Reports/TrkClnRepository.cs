using Cis.Application.Abstractions;
using Cis.Application.Reports;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Reports;

/// <summary><see cref="ITrkClnRepository"/> on cis.TrkCln_*.</summary>
public sealed class TrkClnRepository : ITrkClnRepository
{
    /// <inheritdoc />
    public async Task<(List<string> Groups, List<string> Cities, List<string> States)> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.TrkCln_Lookups");
        var groups = (await grid.ReadAsync<string>()).ToList();
        var cities = (await grid.ReadAsync<string>()).ToList();
        return (groups, cities, (await grid.ReadAsync<string>()).ToList());
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, TrkClnFilter f) =>
        session.RowsAsync("cis.TrkCln_Search", new
        {
            code = f.Code, name = f.Name, group = f.Group, postcode = f.Postcode, city = f.City, state = f.State, contact = f.Contact, doctor = f.Doctor,
            initial = f.Initial, bnfType = f.BnfType, receivedFrom = f.ReceivedFrom, receivedTo = f.ReceivedTo, panel = f.Panel, type = f.Type, xray = f.Xray,
            forRef = f.ForRef, location = f.Location, ffs = f.Ffs,
        }, 300);
}
