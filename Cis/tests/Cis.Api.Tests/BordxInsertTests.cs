using Cis.Application.Bordx;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Bordx;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Insert Worksheet tick rules (no database).</summary>
public class BordxInsertRuleTests
{
    private static Dictionary<string, object?> Row(Action<Dictionary<string, object?>>? change = null)
    {
        var r = new Dictionary<string, object?>
        {
            ["DateVisit"] = new DateTime(2021, 7, 27), ["CLNName"] = "KLINIK", ["MBMGroupCompany"] = "CO", ["OPBatchNo"] = "OP1", ["CLNFFSStatus"] = "N",
            ["CLMMCOProcFee"] = 0m, ["GRPMCOStatus"] = "Y", ["CLMBordxAmt"] = 45m, ["INVPayTo"] = "C", ["CLMTotalActual"] = 45m, ["INVBillAmt"] = 45m, ["CLMCoPayColl"] = 0m,
        };
        change?.Invoke(r);
        return r;
    }

    [Fact]
    public void Tick_rules_follow_the_desktop_order()
    {
        Assert.Null(BordxWsService.TickError(Row(), false));
        Assert.Equal("Date of Visit cannot be empty.", BordxWsService.TickError(Row(r => r["DateVisit"] = null), false));
        Assert.Equal("Batch Number is empty!", BordxWsService.TickError(Row(r => r["OPBatchNo"] = ""), false));
        Assert.Equal("FFS is empty!", BordxWsService.TickError(Row(r => r["CLNFFSStatus"] = "Y"), false));
        Assert.Equal("Please check Bordx Amount!", BordxWsService.TickError(Row(r => r["CLMBordxAmt"] = 0m), false));
        Assert.Equal("Difference in Bill Amount and Payable to Clinic Amount!", BordxWsService.TickError(Row(r => r["INVBillAmt"] = 40m), false));
        Assert.Equal("Please check Bill Amount & Bordx Amount!", BordxWsService.TickError(Row(r => { r["INVPayTo"] = "M"; r["CLMBordxAmt"] = 30m; }), false));
        Assert.Null(BordxWsService.TickError(Row(r => { r["INVPayTo"] = "M"; r["CLMBordxAmt"] = 30m; }), true));
    }
}

/// <summary>Insert Worksheet against the real databases (skipped without CIS_TEST_DB); the bordx is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class BordxInsertDatabaseTests
{
    private static readonly BordxWsRequest M005 = new() { BordxType = 2, GroupCompany = "X|M005" };

    [Fact]
    public async Task Search_checks_and_a_rolled_back_bordx()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BordxWsService(TestDb.Sessions(), new BordxWsRepository());
        Assert.Equal(BordxWsService.NoType, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new BordxWsRequest()))).Message);
        Assert.Equal(BordxWsService.NoAso, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new BordxWsRequest { BordxType = 1 }))).Message);
        Assert.Equal(BordxWsService.NoWorksheet,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(M005 with { Scan = true }))).Message);
        var r = await service.SearchAsync(M005);
        var row = Assert.Single(r.Rows);
        Assert.Equal("EA12573444", ((string)row["CASNumber"]!).Trim());

        var repo = new BordxWsRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        var no = await repo.NextNoAsync(tx, "CN");
        Assert.StartsWith("KCN", no);
        var amount = await repo.UpdateCaseAsync(tx, "EA12573444", no, 0, "DEVADMIN");
        Assert.True(amount > 0);
        await repo.SaveBordxAsync(tx, no, "DC", "DEVADMIN", amount, 1, "X", "", "");
        var status = await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT RTRIM(BordxRefNo) FROM CISWorksheet.dbo.Claim WHERE CASNumber = 'EA12573444'", transaction: tx.Sql().Transaction);
        Assert.Equal(no, status);
    }
}
