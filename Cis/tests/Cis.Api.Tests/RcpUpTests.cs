using Cis.Application.Receipt;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Receipt;
using ClosedXML.Excel;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Upload Receipt against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class RcpUpDatabaseTests
{
    private static MemoryStream Sheet(bool multiple)
    {
        using var book = new XLWorkbook();
        var s = book.AddWorksheet("Sheet1");
        s.Cell(2, 2).Value = "KIT00036";
        s.Cell(2, 3).Value = "EA12523998";
        s.Cell(2, 6).Value = 90;
        s.Cell(2, 7).Value = "PV1";
        s.Cell(4, 2).Value = "KIT00036";
        s.Cell(4, 3).Value = "NOPE";
        s.Cell(4, 6).Value = 5;
        if (multiple)
        {
            s.Cell(2, 8).Value = "RPTESTM";
            s.Cell(2, 9).Value = "MXTEST";
        }
        var ms = new MemoryStream();
        book.SaveAs(ms);
        ms.Position = 0;
        return ms;
    }

    [Fact]
    public async Task Upload_matches_and_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new RcpUpService(TestDb.Sessions(), new RcpUpRepository(), new SpreadsheetReader());
        var rows = await service.UploadAsync(Sheet(false), false);
        Assert.Equal(2, rows.Count);
        Assert.Equal(RcpUpService.Matched, rows[0].CaseStatus);
        Assert.Equal(9m, rows[0].Difference);
        Assert.Equal(RcpUpService.NotMatched, rows[1].CaseStatus);
        Assert.Equal(RcpUpService.NoMedixNo,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new RcpUpSaveRequest { ReceiptNo = "X", Rows = rows }, "DEVADMIN"))).Message);
        var multi = await service.UploadAsync(Sheet(true), true);
        Assert.Equal("RPTESTM", multi[0].ReceiptNo);

        await using var tx = await TestDb.Sessions().BeginAsync();
        await new RcpUpRepository().SaveAsync(tx, "EA12523998", null, "RPTESTM", DateTime.Today, 90m, "CHQ", 90m, "11", "2026", "PV1", "MXTEST", "DEVADMIN");
        Assert.Equal("KIT00036", await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT RTRIM(BordxNo) FROM CISPayment.dbo.ReceiptfrPAY WHERE ReceiptNo = 'RPTESTM'", transaction: tx.Sql().Transaction));
    }
}
