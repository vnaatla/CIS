using Cis.Application.Bordx;
using Cis.Domain.Errors;
using Cis.Infrastructure.Bordx;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Insert Invoice against the real databases (skipped without CIS_TEST_DB); the bordx is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class BordxInvDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_bordx()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BordxInvService(TestDb.Sessions(), new BordxInvRepository(), new BordxWsRepository());
        var result = await service.SearchAsync(new BordxInvSearchRequest("01/08/2021", "04/08/2021", null, "A242 - ASIAFLEX PRODUCTS SDN. BHD."));
        Assert.Contains(result.Rows, x => (string?)x["TickError"] == BordxInvService.NoDate);
        Assert.Contains(result.Rows, x => (x["InvoiceNo"] as string)?.Trim() == "7486" && x["TickError"] is null);
        Assert.Equal(BordxInvService.NoneTicked,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.BordxAsync(new BordxInvSaveRequest(null, []), "DEVADMIN"))).Message);

        var repo = new BordxInvRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        var row = result.Rows.First(x => (x["InvoiceNo"] as string)?.Trim() == "7486");
        var cases = await repo.CasesAsync(tx, new BordxInvTick("7486", (DateTime?)row["InvoiceDate"], (DateTime?)row["RegisterDate"], "A242"));
        var (cas, _) = Assert.Single(cases);
        Assert.Equal("EA12589360", cas);
        var no = await new BordxWsRepository().NextNoAsync(tx, "DC");
        await repo.UpdateCaseAsync(tx, cas, no);
        Assert.Equal(no, await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT RTRIM(BordxRefNo) FROM CISWorksheet.dbo.Claim WHERE CASNumber = 'EA12589360'", transaction: tx.Sql().Transaction));
    }
}
