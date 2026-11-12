using Cis.Application.Abstractions;
using Cis.Application.AcExport;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.AcExport;

/// <summary><see cref="IPostBordxGrpRepository"/> on cis.SunPost_BordxGrp_*.</summary>
public sealed class PostBordxGrpRepository : IPostBordxGrpRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostBordxGrpFilter f) =>
        session.RowsAsync("cis.SunPost_BordxGrp_Search", new
        {
            batchNo = f.BatchNo, company = f.Company, bordxFrom = f.BordxFrom, bordxTo = f.BordxTo, bordxDateFrom = f.BordxDateFrom, bordxDateTo = f.BordxDateTo,
            postDateFrom = f.PostDateFrom, postDateTo = f.PostDateTo, posted = f.Posted, batchFrom = f.BatchFrom, batchTo = f.BatchTo, sj = f.Sj,
        }, 1800);

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, string bordxNo, string accMonth, string accYear, string batchNo, string user) =>
        session.ExecAsync("cis.SunPost_BordxGrp_Insert", new { bordxNo, accMonth, accYear, batchNo, user });
}
