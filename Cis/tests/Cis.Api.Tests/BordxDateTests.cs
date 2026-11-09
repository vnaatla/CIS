using Cis.Application.Bordx;
using Cis.Domain.Errors;
using Cis.Infrastructure.Bordx;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>ClaimBordxDate against the real databases (skipped without CIS_TEST_DB); the update is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class BordxDateDatabaseTests
{
    [Fact]
    public async Task Find_checks_and_a_rolled_back_update()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BordxDateService(TestDb.Sessions(), new BordxDateRepository());
        Assert.Equal(BordxDateService.NotFound, (await Assert.ThrowsAsync<NotFoundException>(() => service.FindAsync("KZZ99999"))).Message);
        var (rows, _) = await service.FindAsync("KIV03270");
        Assert.Equal("KIV03270", Assert.Single(rows).BordxRefNo.Trim());
        Assert.Equal(BordxDateService.NoDate,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.UpdateAsync(new BordxDateUpdateRequest { BordxNo = "KIV03270" }, "DEVADMIN"))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        await new BordxDateRepository().UpdateAsync(tx, "KIV03270", new DateTime(2021, 8, 5), "DEVADMIN");
        var month = await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT TOP 1 CLMBordxMonth FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = 'KIV03270'", transaction: tx.Sql().Transaction);
        Assert.Equal("08", month.Trim());
    }
}
