using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.Entity;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IEntityGrpComAccessRepository"/> on the <c>cis.EntityGrpCom_*</c> / <c>cis.EntityAccess_*</c> procedures.</summary>
public sealed class EntityGrpComAccessRepository : IEntityGrpComAccessRepository
{
    /// <inheritdoc />
    public async Task<GroupCompanyClinicLookups> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.EntityGrpCom_Lookups");
        return new GroupCompanyClinicLookups((await grid.ReadAsync<LookupOption>()).ToList(), (await grid.ReadAsync<LookupOption>()).ToList());
    }

    /// <inheritdoc />
    public async Task<GroupCompanyClinicPage> SearchAsync(IDbSession session, string? grpCom, string? cln, int page, int pageSize, string? sortKey, bool sortDesc)
    {
        using var grid = await session.MultipleAsync("cis.EntityGrpCom_Search", new { grpCom, cln, page, pageSize, sortKey, sortDesc });
        return new GroupCompanyClinicPage((await grid.ReadAsync<GroupCompanyClinic>()).ToList(), await grid.ReadSingleAsync<long>());
    }

    /// <inheritdoc />
    public Task<bool> ExistsAsync(IDbSession session, string grpCom, string cln) => session.ScalarAsync<bool>("cis.EntityGrpCom_Exists", new { grpCom, cln });

    /// <inheritdoc />
    public Task<int> SaveAsync(IDbSession session, bool insert, string? oldGrpCom, string? oldCln, string grpCom, string cln, string userCode, DateTime today) =>
        session.ScalarAsync<int>("cis.EntityGrpCom_Save", new { insert, oldGrpCom, oldCln, grpCom, cln, user = userCode, today });

    /// <inheritdoc />
    public Task<int> DeleteAsync(IDbSession session, string grpCom, string cln) => session.ScalarAsync<int>("cis.EntityGrpCom_Delete", new { grpCom, cln });

    /// <inheritdoc />
    public Task<List<ClinicAccess>> AccessListAsync(IDbSession session) => session.QueryAsync<ClinicAccess>("cis.EntityAccess_List");

    /// <inheritdoc />
    public Task<string?> ClinicNameAsync(IDbSession session, string code) => session.FirstOrDefaultAsync<string>("cis.EntityAccess_ClinicName", new { code });

    /// <inheritdoc />
    public Task InsertAccessAsync(IDbSession session, string code, DateTime from, DateTime to, string userCode, DateTime now) =>
        session.ExecAsync("cis.EntityAccess_Insert", new { code, from, to, user = userCode, now });

    /// <inheritdoc />
    public Task<List<ClinicFound>> FindClinicsAsync(IDbSession session, string name) => session.QueryAsync<ClinicFound>("cis.EntityAccess_FindClinics", new { name });
}
