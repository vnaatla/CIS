using Cis.Application.Mco;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Mco;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Fee Listing against the real databases (skipped without CIS_TEST_DB); the invoice print is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class McoFeeListingDatabaseTests
{
    private static McoListService Service(Cis.Application.Abstractions.IDbSessions sessions)
    {
        PdfSharp.Fonts.GlobalFontSettings.FontResolver ??= new PdfFontResolver();
        return new(sessions, new McoListRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()), TimeProvider.System);
    }

    private static readonly McoListRequest Invoiced = new() { Criteria = 2, InvoiceNo = "I07600", Version = 1, GroupCompany = "M284 - X" };

    [Fact]
    public async Task Search_checks_and_reports()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = Service(TestDb.Sessions());
        Assert.Equal(McoListService.NoCriteria, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new McoListRequest()))).Message);
        Assert.Equal(McoListService.NoBatch, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new McoListRequest { Criteria = 0 }))).Message);
        Assert.Equal(McoListService.NoGroup,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new McoListRequest { Criteria = 2, InvoiceNo = "I1", Version = 1 }))).Message);

        var r = await service.SearchAsync(Invoiced);
        Assert.Equal(5363, r.Rows.Count);
        Assert.All(r.Rows, x => Assert.Equal("M284", ((string)x["GRPCoCode"]!).Trim()));
        Assert.False(r.CanPrint);
        Assert.False(r.MultipleInvoices);
        var listing = await service.PrintAsync(new McoListPrintRequest { Search = Invoiced, ReportType = 1 }, "DEVADMIN");
        Assert.Equal("application/pdf", listing.ContentType);
        var invoice = await service.PrintAsync(new McoListPrintRequest { Search = Invoiced, ReportType = 0, Print = true }, "DEVADMIN");
        Assert.Equal("MCOInvoice_I07600.pdf", invoice.FileName);
        Assert.Equal(McoListService.NoReportType,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PrintAsync(new McoListPrintRequest { Search = Invoiced }, "DEVADMIN"))).Message);
    }

    [Fact]
    public async Task Printing_a_new_invoice_numbers_and_stamps_it_rolled_back()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new McoListRepository();
        await using var s = await TestDb.Sessions().BeginAsync();
        var before = await s.Sql().Connection.QuerySingleAsync<string>("SELECT TOP 1 MCOInvoice_No FROM dbo.CLNMCOINVQ", transaction: s.Sql().Transaction);
        var no = await repo.NextInvoiceNoAsync(s);
        Assert.Equal("I" + (int.Parse(before) + 1).ToString("00000"), no);
        var rows = await repo.SearchAsync(s, new McoListFilter(0, new DateTime(2021, 8, 6), 1, null, null, null, "N", null, "D005"));
        Assert.NotEmpty(rows);
        var row = rows[0];
        await repo.StampAsync(s, (string)row["MBMNumber"]!, ((string)row["MBMCCoverID"]!).Trim(), no, "DEVADMIN", 9.5m);
        var after = await repo.SearchAsync(s, new McoListFilter(2, null, null, null, null, no, "N", null, null));
        Assert.Single(after);
    }
}
