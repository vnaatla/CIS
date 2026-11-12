using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Reports;

namespace Cis.Api.Tests;

/// <summary>Tracking > Grp. Com. Clinic against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class TrkGrpClnDatabaseTests
{
    [Fact]
    public async Task The_four_views_and_checks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new TrkGrpClnService(TestDb.Sessions(), new TrkGrpClnRepository());
        var (companies, clinics) = await service.LookupsAsync();
        Assert.NotEmpty(companies);
        Assert.Contains("KLINIK BALA", clinics);
        var perClinic = Assert.Single(await service.SearchAsync(new TrkGrpClnRequest(4, null, "B324", null, "KLINIK BALA")));
        Assert.True(Convert.ToInt32(perClinic["TotalCom"]) > 0);
        Assert.NotEmpty(await service.SearchAsync(new TrkGrpClnRequest(1, null, "B324", null, null)));
        Assert.Equal(TrkGrpClnService.NoRecord,
            (await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(new TrkGrpClnRequest(2, "ZZZZZZ", null, null, null)))).Message);
    }
}
