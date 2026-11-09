using Cis.Application.Bordx;
using Cis.Domain.Errors;
using Cis.Infrastructure.Bordx;
using Cis.Infrastructure.Documents;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Print Bordx against the real databases (skipped without CIS_TEST_DB); previews only (nothing is stamped).</summary>
[Collection(DatabaseCollection.Name)]
public class PrintBordxDatabaseTests
{
    [Fact]
    public async Task Find_checks_and_preview()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PrintBordxService(TestDb.Sessions(), new PrintBordxRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()));
        Assert.Equal(PrintBordxService.NoBordx, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.FindAsync(""))).Message);
        var info = await service.FindAsync("KDC55049");
        Assert.False(info.Printed);
        Assert.Equal(250m, info.TotalAmount);
        var req = new PrintBordxRequest { BordxNo = "KDC55049", Company = "Health Connect", BillType = 1, ReportType = 1 };
        Assert.Equal(PrintBordxService.NoMonth, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PrintAsync(req, "DEVADMIN"))).Message);
        Assert.Equal(PrintBordxService.NoBillType,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PrintAsync(req with { BillType = null }, "DEVADMIN"))).Message);
        var file = await service.PrintAsync(req with { Month = "JUNE", Year = "2021" }, "DEVADMIN");
        using var book = new ClosedXML.Excel.XLWorkbook(new MemoryStream(file.Content));
        Assert.Equal(2, book.Worksheets.Count);
        foreach (var v in new[] { "seb", "prasarana", "axa-affin", "time-dotcom", "bsn", "berjaya-sompo", "crystal" })
        {
            var f = await service.PrintAsync(new PrintBordxRequest { BordxNo = "KDC55049", Variant = v, Month = "June", Year = "2021", Summary = true }, "DEVADMIN");
            Assert.EndsWith(".xlsx", f.FileName);
        }
    }
}
