using Cis.Application.Abstractions;
using Cis.Application.Maintenance.Entity;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IEntityGroupDoctorRepository"/> on the <c>cis.EntityGroup_*</c> / <c>cis.EntityDoctor_*</c> procedures.</summary>
public sealed class EntityGroupDoctorRepository : IEntityGroupDoctorRepository
{
    /// <inheritdoc />
    public async Task<List<EntityListRow>> SearchGroupsAsync(IDbSession session, string? code, string? name) =>
        (await session.QueryAsync<(string GRPClnCode, string? GRPClnName)>("cis.EntityGroup_Search", new { code, name }))
        .Select(r => new EntityListRow(r.GRPClnCode, r.GRPClnName)).ToList();

    /// <inheritdoc />
    public Task<ClinicGroupDto?> GetGroupAsync(IDbSession session, string code) => session.FirstOrDefaultAsync<ClinicGroupDto>("cis.EntityGroup_Get", new { code });

    /// <inheritdoc />
    public async Task<string> NextGroupCodeAsync(IDbSession session, string letter) =>
        (await session.ScalarAsync<string>("cis.EntityGroup_NextCode", new { letter }))!;

    /// <inheritdoc />
    public Task<int> SaveGroupAsync(IDbSession session, bool insert, ClinicGroupDto g, string userCode, DateTime today) =>
        session.ScalarAsync<int>("cis.EntityGroup_Save", new
        {
            insert, code = g.Code, name = g.Name, initial = g.Initial, add1 = g.Add1, add2 = g.Add2, add3 = g.Add3, contact = g.Contact,
            admContact = g.AdmContact, ceo = g.Ceo, phone = g.Phone, fax = g.Fax, email = g.Email, homepage = g.Homepage,
            effDate = g.EffDate is null ? (DateTime?)null : DateTime.Parse(g.EffDate, System.Globalization.CultureInfo.InvariantCulture),
            tax = g.Tax, conversion = g.Conversion, prePost = g.PrePost, user = userCode, today,
        });

    /// <inheritdoc />
    public Task<bool> GroupInUseAsync(IDbSession session, string code) => session.ScalarAsync<bool>("cis.EntityGroup_InUse", new { code });

    /// <inheritdoc />
    public Task<int> DeleteGroupAsync(IDbSession session, string code) => session.ScalarAsync<int>("cis.EntityGroup_Delete", new { code });

    /// <inheritdoc />
    public async Task<List<EntityListRow>> SearchDoctorsAsync(IDbSession session, string? code, string? name) =>
        (await session.QueryAsync<(string DocCode, string? DocName)>("cis.EntityDoctor_Search", new { code, name }))
        .Select(r => new EntityListRow(r.DocCode, r.DocName)).ToList();

    /// <inheritdoc />
    public Task<bool> DoctorNameExistsAsync(IDbSession session, string name) => session.ScalarAsync<bool>("cis.EntityDoctor_NameExists", new { name });

    /// <inheritdoc />
    public async Task<string> NextDoctorCodeAsync(IDbSession session, string letter) =>
        (await session.ScalarAsync<string>("cis.EntityDoctor_NextCode", new { letter }))!;

    /// <inheritdoc />
    public Task<int> SaveDoctorAsync(IDbSession session, bool insert, string code, string name, string userCode, DateTime today) =>
        session.ScalarAsync<int>("cis.EntityDoctor_Save", new { insert, code, name, user = userCode, today });

    /// <inheritdoc />
    public Task<int> DeleteDoctorAsync(IDbSession session, string code) => session.ScalarAsync<int>("cis.EntityDoctor_Delete", new { code });
}
