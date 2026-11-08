using Cis.Application.Mco;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Mco;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>MCO Invoice against the real databases (skipped without CIS_TEST_DB); stamps are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class McoInvoiceDatabaseTests
{
    private static readonly McoInvoiceRequest D048 = new() { Criteria = 0, BordxDate = "06/08/2021", BatchNo = "1", Version = 1, GroupCompany = "D048 - X", PrintOption = 0 };

    private static McoInvoiceService Service() =>
        new(TestDb.Sessions(), new McoInvoiceRepository(), new McoListRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()), TimeProvider.System);

    [Fact]
    public async Task Search_checks_lists_and_previews()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = Service();
        Assert.Equal(McoInvoiceService.NoBordx, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new McoInvoiceRequest { Criteria = 0 }))).Message);
        Assert.Equal(McoInvoiceService.NoCompany, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(D048 with { GroupCompany = "" }))).Message);
        var r = await service.SearchAsync(D048);
        Assert.Equal(2, r.Principals);
        Assert.Equal(4, r.Supplementaries);
        Assert.Equal("", r.InvoiceNo);
        var file = await service.PrintAsync(D048, "DEVADMIN");
        Assert.EndsWith(".xlsx", file.FileName);
    }

    [Fact]
    public async Task Summary_lines_and_a_rolled_back_stamp()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new McoInvoiceRepository();
        await using var s = await TestDb.Sessions().BeginAsync();
        var f = new McoInvoiceFilter(0, new DateTime(2021, 8, 6), 1, null, null, null, 1, null, "D048", null, false);
        var lines = McoInvoiceService.Summary(await repo.SummaryAsync(s, f));
        var line = Assert.Single(lines);
        Assert.Equal(("H - GROUP FAMILY", "Child", 2m, 6m), (line["Type"], line["AdultChild"], line["Principals"], line["Covered"]));
        await repo.StampAsync(s, "STHB9440145*01", "00", "IZZTEST", "DEVADMIN");
        var after = await repo.PrincipalsAsync(s, f with { Mode = 2, InvoiceNo = "IZZTEST" });
        Assert.Single(after);
    }
}
