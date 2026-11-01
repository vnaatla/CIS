using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.Entity;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IEntityClinicRepository"/> on the <c>cis.EntityClinic_*</c> procedures.</summary>
public sealed class EntityClinicRepository : IEntityClinicRepository
{
    /// <inheritdoc />
    public async Task<ClinicSetup> SetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.EntityClinic_Setup");
        return new ClinicSetup((await grid.ReadAsync<LookupOption>()).ToList(), (await grid.ReadAsync<LookupOption>()).ToList(),
            (await grid.ReadAsync<LookupOption>()).ToList());
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, ClinicDto f) =>
        session.RowsAsync("cis.EntityClinic_Search", new
        {
            code = f.Code, initial = f.Initial, name = f.Name, state = f.State, city = f.City, postCode = f.PostCode, panel = f.Panel,
            groupPrefix = f.GroupCode, type = f.Type, hrs24 = f.Hrs24, ffs = f.Ffs, online = f.Online, glFormat = f.GlFormat,
        });

    /// <inheritdoc />
    public Task<bool> AddressExistsAsync(IDbSession session, string add1) => session.ScalarAsync<bool>("cis.EntityClinic_AddressExists", new { add1 });

    /// <inheritdoc />
    public async Task<string> NextCodeAsync(IDbSession session, string letter) =>
        (await session.ScalarAsync<string>("cis.EntityClinic_NextCode", new { letter }))!;

    /// <inheritdoc />
    public Task<int> SaveAsync(IDbSession session, bool insert, ClinicDto c, string userCode, DateTime today) =>
        session.ScalarAsync<int>("cis.EntityClinic_Save", new
        {
            insert, code = c.Code, name = c.Name, groupCode = c.GroupCode, phone = c.Phone, fax = c.Fax, add1 = c.Add1, add2 = c.Add2, add3 = c.Add3,
            email = c.Email, homepage = c.Homepage, representative = c.Contact, docName = c.DocName,
            bnfType = int.TryParse(c.BnfType, out var b) ? b : (int?)null, forRef = c.FormemaRef,
            noOfDoc = int.TryParse(c.NoOfDoc, out var n) ? n : (int?)null, type = c.Type, xray = c.XRay, panel = c.Panel, hrs24 = c.Hrs24, ffs = c.Ffs,
            ffsPerc = decimal.TryParse(c.FfsPerc, NumberStyles.Number, CultureInfo.InvariantCulture, out var p) ? p : (decimal?)null,
            state = c.State, city = c.City, postCode = c.PostCode, initial = c.Initial,
            receivedDate = c.ReceivedDate is null ? (DateTime?)null : DateTime.Parse(c.ReceivedDate, CultureInfo.InvariantCulture),
            bankAcc = c.BankAcc, bankName = c.Bank, bankBrn = c.Branch, accPerson = c.AccPerson, remark = c.Remark, location = c.Location,
            operatingHr = c.OperatingHr, pName = c.PayName, pAdd1 = c.PayAdd1, pAdd2 = c.PayAdd2, pAdd3 = c.PayAdd3, online = c.Online,
            glFormat = c.GlFormat, regNo = c.RegNo, user = userCode, today,
        });

    /// <inheritdoc />
    public Task<bool> HasCasesAsync(IDbSession session, string code) => session.ScalarAsync<bool>("cis.EntityClinic_HasCases", new { code });

    /// <inheritdoc />
    public Task<int> DeleteAsync(IDbSession session, string code) => session.ScalarAsync<int>("cis.EntityClinic_Delete", new { code });
}
