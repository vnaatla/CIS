using Cis.Application.Bordx;
using Cis.Domain.Errors;
using Cis.Infrastructure.Bordx;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Remove Worksheet against the real databases (skipped without CIS_TEST_DB); the removal is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class BordxRemoveDatabaseTests
{
    [Fact]
    public async Task Find_and_a_rolled_back_removal()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BordxRemoveService(TestDb.Sessions(), new BordxRemoveRepository());
        Assert.Equal(BordxRemoveService.NoBordx, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.FindAsync(" "))).Message);
        var r = await service.FindAsync("KIV03270");
        Assert.Contains(r.Rows, x => ((string)x["CASNumber"]!).Trim() == "EA12476058");

        var repo = new BordxRemoveRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        var before = await tx.Sql().Connection.QuerySingleAsync<(short Claims, decimal Amount)>(
            "SELECT BordxNoClaims, BordxClaimAmt FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = 'KIV03270'", transaction: tx.Sql().Transaction);
        await repo.RemoveAsync(tx, "KIV03270", "EA12476058", "DEVADMIN");
        var after = await tx.Sql().Connection.QuerySingleAsync<(short Claims, decimal Amount)>(
            "SELECT BordxNoClaims, BordxClaimAmt FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = 'KIV03270'", transaction: tx.Sql().Transaction);
        Assert.Equal(before.Claims - 1, after.Claims);
        Assert.True(after.Amount < before.Amount);
        Assert.DoesNotContain(await repo.ListAsync(tx, "KIV03270"), x => ((string)x["CASNumber"]!).Trim() == "EA12476058");
    }
}
