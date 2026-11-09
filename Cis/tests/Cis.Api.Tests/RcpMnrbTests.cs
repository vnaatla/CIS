using Cis.Application.Receipt;
using Cis.Domain.Errors;
using Cis.Infrastructure.Receipt;

namespace Cis.Api.Tests;

/// <summary>Receipt from MNRB rules (no database).</summary>
public class RcpMnrbRuleTests
{
    [Fact]
    public void Shape_follows_the_desktop_lists()
    {
        var partial = new Dictionary<string, object?> { ["BordxMNRBFRStatus"] = "P", ["BordxClaimAmt"] = 100m, ["CheckAmt"] = 40m };
        RcpMnrbService.Shape(partial, false);
        Assert.Equal("Partial", partial["Status"]);
        Assert.Equal(60m, partial["TickAmount"]);
        var open = new Dictionary<string, object?> { ["BordxMNRBFRStatus"] = null, ["BordxClaimAmt"] = 100m, ["CheckAmt"] = null };
        RcpMnrbService.Shape(open, false);
        Assert.Equal(100m, open["Balance"]);
    }
}

/// <summary>Receipt from MNRB against the real databases (skipped without CIS_TEST_DB). The restored data has no MNRB bordereaux.</summary>
[Collection(DatabaseCollection.Name)]
public class RcpMnrbDatabaseTests
{
    [Fact]
    public async Task Payors_search_and_checks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new RcpMnrbService(TestDb.Sessions(), new RcpMnrbRepository());
        Assert.Contains(await service.PayorsAsync(), p => p.Code == "JE");
        Assert.Equal(RcpMnrbService.NoBordx, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new RcpMnrbSearchRequest()))).Message);
        Assert.Equal(RcpMnrbService.NoRecordWks,
            (await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(new RcpMnrbSearchRequest { Worksheet = true, Scan = false }))).Message);
        Assert.Equal(RcpMnrbService.NoneSelected, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new RcpMnrbSaveRequest(), "DEVADMIN"))).Message);
        Assert.Equal(RcpMnrbService.NotPaid, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.ReceiptAsync("NOPE", null))).Message);
    }
}
