using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Reports;

namespace Cis.Api.Tests;

/// <summary>Tracking > Clinic against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class TrkClnDatabaseTests
{
    [Fact]
    public async Task Lookups_search_and_checks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new TrkClnService(TestDb.Sessions(), new TrkClnRepository());
        var (groups, cities, states) = await service.LookupsAsync();
        Assert.NotEmpty(groups);
        Assert.NotEmpty(cities);
        Assert.NotEmpty(states);
        var rows = await service.SearchAsync(new TrkClnRequest { Name = "BALA" });
        Assert.Contains(rows, r => (r["CLNCode"] as string)?.Trim() == "B020");
        Assert.Equal(TrkClnService.NoRecord, (await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(new TrkClnRequest { Code = "ZZZZ" }))).Message);
        Assert.Equal(TrkClnService.BadDate,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new TrkClnRequest { ReceivedFrom = "31/02/2021" }))).Message);
    }
}
