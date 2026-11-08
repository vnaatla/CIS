using Cis.Application.Mco;
using Cis.Domain.Errors;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Mco;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>GST Zeroised against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class McoZeroDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_the_zeroised_invoice()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new McoZeroService(TestDb.Sessions(), new McoZeroRepository(), new McoInvoiceRepository(), new McoListRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()), TimeProvider.System);
        var batch = new McoZeroRequest { Search = new McoGstRequest { Criteria = 0, BordxDate = "06/08/2021", BatchNo = "1", Version = 1, GroupCompany = "D005 - X" } };
        Assert.Equal(6, (await service.SearchAsync(batch)).Rows.Count);
        Assert.Equal(McoZeroService.NoChecked, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PrintAsync(batch, "DEVADMIN"))).Message);
        Assert.Equal(McoZeroService.NoInvoiceSearch, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PrintInvoiceAsync(batch))).Message);
        var file = await service.PrintInvoiceAsync(new McoZeroRequest { Search = new McoGstRequest { Criteria = 2, InvoiceNo = "I07600" } });
        using var book = new ClosedXML.Excel.XLWorkbook(new MemoryStream(file.Content));
        Assert.Equal(2, book.Worksheets.Count);
        Assert.StartsWith("I07600-", file.FileName);
    }
}
