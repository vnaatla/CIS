using Cis.Application.Receipt;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Receipt;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Receipt from Group Company against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class RcpPayDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new RcpPayService(TestDb.Sessions(), new RcpPayRepository());
        Assert.Equal("DC", RcpPayService.PayorCode("DC - DUMMY"));
        Assert.Equal(RcpPayService.NoBordx, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new RcpPaySearchRequest()))).Message);
        var bordx = Assert.Single(await service.SearchAsync(new RcpPaySearchRequest { BordxFrom = "KIT00036" }));
        Assert.Equal("Full", bordx["Status"]);
        Assert.Equal("RP37929", ((string)bordx["ReceiptNo"]!).Trim());
        var wks = Assert.Single(await service.SearchAsync(new RcpPaySearchRequest { Worksheet = true, CaseFrom = "EA12523998" }));
        Assert.Equal(RcpPayService.Paid, wks["TickError"]);
        Assert.Equal(RcpPayService.NoReceiptNo, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new RcpPaySaveRequest(), "DEVADMIN"))).Message);
        Assert.Equal(RcpPayService.Paid, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new RcpPaySaveRequest
        {
            ReceiptNo = "RPTEST", ReceiptDate = "01/11/2026", ChequeNo = "CHQ", ReceiptAmt = "99", Ticked = [new("KIT00036", null)],
        }, "DEVADMIN"))).Message);

        var repo = new RcpPayRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        await repo.SummaryAsync(tx, "RPTEST", DateTime.Today, 1m, "OPCLM", null, "CHQ", "X");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISPayment.dbo.ReceiptSummary WHERE ReceiptNo = 'RPTEST'", transaction: tx.Sql().Transaction));
    }
}
