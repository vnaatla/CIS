using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Reports;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Email Notification (Providers) rules (no database).</summary>
public class EmlPrvRuleTests
{
    [Fact]
    public void The_body_has_the_invoice_table_and_the_entity()
    {
        var d = new Dictionary<string, object?>
        {
            ["InvoiceNo"] = "INV1", ["InvoiceDate"] = new DateTime(2021, 7, 1), ["PVCLNDate"] = new DateTime(2021, 8, 3), ["CLMTotalActual"] = 400m,
            ["CLMCoPayColl"] = 10m, ["CLMMCOProcFee"] = 2.1m, ["CLMPayableByMedix2c"] = 387.9m, ["ENTITYCODE"] = "SJ",
        };
        var (entity, html) = EmlPrvService.Body(new EmlPrvTick("CP1", "a@b.c", "A033", "KLINIK", "CHQ1", "387.9"), [d], "05/08/2021");
        Assert.Equal("MAJU SIHAT", entity);
        Assert.Contains("<tr><td>1</td> <td>INV1</td> <td>01/07/2021</td> <td>03/08/2021</td> <td>400.00</td> <td>10.00</td> <td>2.10</td> <td>387.90</td> </tr>", html);
        Assert.Contains("PV Amount               :       RM 387.90</p>", html);
        Assert.EndsWith("Regards,</p>MAJU SIHAT SDN BHD </BODY></HTML>", html);
    }
}

/// <summary>Email Notification (Providers) against the real databases (skipped without CIS_TEST_DB); the mark is rolled back and nothing is e-mailed.</summary>
[Collection(DatabaseCollection.Name)]
public class EmlPrvDatabaseTests
{
    [Fact]
    public async Task Search_invoices_and_rolled_back_mark()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new EmlPrvService(TestDb.Sessions(), new EmlPrvRepository(), null!);
        var rows = await service.SearchAsync(new EmlPrvRequest("03/08/2021", "03/08/2021", "CP100471757", "CP100471757", "N - New", 0));
        Assert.Equal("A033", ((string)Assert.Single(rows)["ClinicCode"]!).Trim());
        Assert.Equal(EmlPayService.NoRecord, (await Assert.ThrowsAsync<NotFoundException>(() =>
            service.SearchAsync(new EmlPrvRequest("03/08/2021", "03/08/2021", "ZZ", "ZZ", "R - Resend", 0)))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        Assert.NotEmpty(await new EmlPrvRepository().InvoicesAsync(tx, "CP100471757"));
        await new EmlPrvRepository().MarkAsync(tx, "CP100471757");
        Assert.True(await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISPayment.dbo.Payment2CLN WHERE PVCLNNo = 'CP100471757' AND PVEmailStatus = 'Y'", transaction: tx.Sql().Transaction) > 0);
    }
}
