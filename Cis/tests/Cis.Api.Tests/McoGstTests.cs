using Cis.Application.Mco;
using Cis.Domain.Errors;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Mco;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>MCO Invoice GST against the real databases (skipped without CIS_TEST_DB); read-only (preview).</summary>
[Collection(DatabaseCollection.Name)]
public class McoGstDatabaseTests
{
    [Fact]
    public async Task Search_and_preview_with_the_detail_sheet()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new McoGstService(TestDb.Sessions(), new McoGstRepository(), new McoInvoiceRepository(), new McoListRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()));
        var req = new McoGstRequest { Criteria = 0, BordxDate = "06/08/2021", BatchNo = "1", Version = 1, GroupCompany = "D005 - X" };
        Assert.Equal(McoInvoiceService.NoCompany, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(req with { GroupCompany = "" }))).Message);
        var r = await service.SearchAsync(req);
        Assert.Equal(6, r.Rows.Count);
        var file = await service.PrintAsync(req, "DEVADMIN");
        using var book = new ClosedXML.Excel.XLWorkbook(new MemoryStream(file.Content));
        Assert.Equal(2, book.Worksheets.Count);
    }
}
