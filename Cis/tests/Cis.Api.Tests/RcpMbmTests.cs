using Cis.Application.Receipt;
using Cis.Domain.Errors;
using Cis.Infrastructure.Receipt;

namespace Cis.Api.Tests;

/// <summary>Receipt from Member rules (no database).</summary>
public class RcpMbmRuleTests
{
    [Fact]
    public void Shape_follows_the_desktop_list()
    {
        var unpaid = new Dictionary<string, object?> { ["PVRecMBMStatus"] = null, ["CLMPayableByMedix2M"] = -50m };
        RcpMbmService.Shape(unpaid);
        Assert.Equal(50m, unpaid["AmountToReceive"]);
        var partial = new Dictionary<string, object?> { ["PVRecMBMStatus"] = "P", ["CLMPayableByMedix2M"] = -50m, ["ReceivedTotal"] = 20m };
        RcpMbmService.Shape(partial);
        Assert.Equal(30m, partial["Balance"]);
        var paid = new Dictionary<string, object?> { ["PVRecMBMStatus"] = "F", ["CLMPayableByMedix2M"] = -50m };
        RcpMbmService.Shape(paid);
        Assert.Equal(RcpMbmService.Paid, paid["TickError"]);
    }
}

/// <summary>Receipt from Member against the real databases (skipped without CIS_TEST_DB). The restored data has no member debit notes.</summary>
[Collection(DatabaseCollection.Name)]
public class RcpMbmDatabaseTests
{
    [Fact]
    public async Task Search_and_checks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new RcpMbmService(TestDb.Sessions(), new RcpMbmRepository());
        Assert.Equal(RcpMbmService.NoWorksheet, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new RcpMbmSearchRequest()))).Message);
        Assert.Equal(RcpMbmService.NoRecord,
            (await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(new RcpMbmSearchRequest { Scan = false, Payment = 2 }))).Message);
        Assert.Equal(RcpMbmService.NoReceiptNo, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new RcpMbmSaveRequest(), "DEVADMIN"))).Message);
        Assert.Equal(RcpMbmService.NotPaid, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.ReceiptAsync("0"))).Message);
    }
}
