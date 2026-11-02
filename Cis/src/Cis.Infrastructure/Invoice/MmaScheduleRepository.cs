using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IMmaScheduleRepository"/> on <c>cis.Mma_Search</c>.</summary>
public sealed class MmaScheduleRepository : IMmaScheduleRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, MmaFilters f) =>
        session.RowsAsync("cis.Mma_Search", new
        {
            mainBp = f.MainBp, opcs = f.Opcs, procedure = f.Procedure, surgCat = f.SurgCat, anaesCat = f.AnaesCat, part = f.Part, pageFrom = f.PageFrom,
            pageTo = f.PageTo, cptCode = f.CptCode, cptDesc = f.CptDesc,
        });
}
