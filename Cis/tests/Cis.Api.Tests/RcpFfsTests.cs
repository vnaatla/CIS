using Cis.Application.Receipt;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Receipt;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Receipt from FFS against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class RcpFfsDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new RcpFfsService(TestDb.Sessions(), new RcpFfsRepository());
        Assert.Equal(RcpFfsService.NoBordx, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new RcpFfsSearchRequest()))).Message);
        var received = Assert.Single((await service.SearchAsync(new RcpFfsSearchRequest { BordxFrom = "KDC01078" })).Rows);
        Assert.Equal(RcpFfsService.Received, received["TickError"]);
        Assert.Equal(9, (await service.SearchAsync(new RcpFfsSearchRequest(), "MR10245")).Rows.Count);
        var open = Assert.Single((await service.SearchAsync(new RcpFfsSearchRequest { BordxFrom = "KDC55710" })).Rows);
        Assert.Equal("", open["Status"]);
        Assert.Equal(RcpFfsService.WrongAmount, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new RcpFfsSaveRequest
        {
            ReceiptNo = "MRTEST", ReceiptDate = "01/11/2026", ReceiptAmt = "1", ChequeNo = "CHQ", Ticked = ["KDC55710"],
        }, "DEVADMIN"))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        await new RcpFfsRepository().SaveAsync(tx, null, "KDC55710", "MRTEST", DateTime.Today, "CHQ", 14.2m, null, "11", "2026", "DEVADMIN");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISPayment.dbo.ReceiptFrFFS WHERE ReceiptNo = 'MRTEST'", transaction: tx.Sql().Transaction));
    }
}
