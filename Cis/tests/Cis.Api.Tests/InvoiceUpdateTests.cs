using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Invoice;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

public class InvoiceUpdateServiceTests
{
    [Fact]
    public async Task Update_checks_follow_the_desktop_order()
    {
        var service = new InvoiceUpdateService(new FakeSessions(), null!, null!, new FixedClock(new DateTime(2026, 11, 2)));
        async Task<string> Msg(InvUpdateRequest r) => (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(r, "DEV"))).Message;
        Assert.Equal(InvoiceUpdateService.SelectRecord, await Msg(new InvUpdateRequest([], "INV1", "31/07/2021")));
        Assert.Equal(InvoiceUpdateService.InvalidDate, await Msg(new InvUpdateRequest(["K1"], "INV1", "__/__/____")));
        Assert.Equal(InvoiceUpdateService.InvoiceFirst, await Msg(new InvUpdateRequest(["K1"], "", "31/07/2021")));
    }

    [Fact]
    public void Filters_are_trimmed_and_the_clinic_code_is_taken_before_the_dash()
    {
        var q = InvoiceUpdateService.Query(new InvUpdateFilters(" K1 ", "", "M813 - KLINIK ERMA", " ", "I00496", "01/07/2021", "31/02/2021", true,
            null, "31/07/2021", "10", "abc"));
        Assert.Equal(new InvUpdateQuery("K1", null, "M813", null, "I00496", new DateTime(2021, 7, 1), null, true, null, new DateTime(2021, 7, 31), 10m, null), q);
    }
}

/// <summary>Invoice Update against the real databases (skipped without CIS_TEST_DB); the update is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class InvoiceUpdateDatabaseTests
{
    private const string Invoice = "I00496-ERMA-M813";

    [Fact]
    public async Task Search_export_and_update()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new InvoiceUpdateRepository();
        var service = new InvoiceUpdateService(TestDb.Sessions(), repo, new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()),
            new FixedClock(new DateTime(2026, 11, 2)));
        var setup = await service.SetupAsync();
        Assert.NotEmpty(setup.Clinics);
        Assert.NotEmpty(setup.GroupCompanies);
        var filters = new InvUpdateFilters(null, null, "M813 - X", null, Invoice, null, null, false, null, null, null, null);
        var rows = await service.SearchAsync(filters);
        Assert.Contains(rows, r => (string?)r["CASNumber"] == "EA12582360");
        Assert.All(rows, r => Assert.Equal(Invoice, ((string?)r["BillNo"])?.Trim()));
        Assert.True((await service.ExportAsync(filters)).Content.Length > 0);
        await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(filters with { InvoiceNo = "ZZ-NO-SUCH-INVOICE" }));

        await using var s = await TestDb.Sessions().BeginAsync();
        Assert.Equal(1, await repo.SaveAsync(s, "EA12582360", "ZZTEST-INV", new DateTime(2026, 10, 31), new DateTime(2026, 11, 2), "DEV"));
        var row = await s.Sql().Connection.QuerySingleAsync<(string InvoiceNo, DateTime InvoiceDate, string RegisterBy)>(
            "SELECT InvoiceNo, InvoiceDate, RegisterBy FROM dbo.CISCAS WHERE CASNumber = 'EA12582360'", transaction: s.Sql().Transaction);
        Assert.Equal(("ZZTEST-INV", new DateTime(2026, 10, 31), "DEV"), row);
    }
}
