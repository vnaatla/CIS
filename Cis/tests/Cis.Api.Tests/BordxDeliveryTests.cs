using Cis.Application.Bordx;
using Cis.Domain.Errors;
using Cis.Infrastructure.Bordx;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Bordx Delivery against the real databases (skipped without CIS_TEST_DB); updates are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class BordxDeliveryDatabaseTests
{
    private static readonly BordxDeliveryRequest Dayang = new() { Company = "DAYANG ENTERPRISE SDN BHD", DateFrom = "01/01/2021", Status = 0 };

    [Fact]
    public async Task Search_checks_and_rolled_back_updates()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BordxDeliveryService(TestDb.Sessions(), new BordxDeliveryRepository());
        Assert.Equal(BordxDeliveryService.NoCompany,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync("payor", new BordxDeliveryRequest { Scan = true }))).Message);
        Assert.Equal(BordxDeliveryService.NoneTicked,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.UpdateAsync("payor", new BordxDeliveryUpdateRequest { Date = "01/09/2021" }))).Message);
        var r = await service.SearchAsync("payor", Dayang);
        Assert.Equal(17, r.Rows.Count);
        var no = ((string)r.Rows[0]["BordxRefNo"]!).Trim();

        var repo = new BordxDeliveryRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        Assert.Equal(1, await repo.UpdateAsync(tx, "payor", no, new DateTime(2021, 9, 1)));
        Assert.Equal(2, await repo.UpdateAsync(tx, "returned", no, new DateTime(2021, 8, 1)));
        Assert.Equal(1, await repo.UpdateAsync(tx, "returned", no, new DateTime(2021, 9, 2)));
        var status = await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT CAST(BordxPayDocFRStatus AS int) FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @no", new { no }, tx.Sql().Transaction);
        Assert.Equal(1, status);
        Assert.Equal(1, await repo.UpdateAsync(tx, "acperiod", no, null, "9", "2021"));
        var period = await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT RTRIM(PAYToACPeriodMth) + '/' + RTRIM(PAYToACPeriodYR) FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @no", new { no }, tx.Sql().Transaction);
        Assert.Equal("9/2021", period);
    }

    [Fact]
    public async Task Ac_period_search_and_checks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BordxDeliveryService(TestDb.Sessions(), new BordxDeliveryRepository());
        Assert.Equal(17, (await service.SearchAsync("acperiod", Dayang with { Status = null })).Rows.Count);
        Assert.Equal(BordxDeliveryService.NoMonth,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.UpdateAsync("acperiod", new BordxDeliveryUpdateRequest { BordxNos = ["X"] }))).Message);
        Assert.Equal(("9", null), BordxDeliveryService.Period("09/2021"));
        Assert.Equal(("11", "2021"), BordxDeliveryService.Period("11/2021"));
    }
}
