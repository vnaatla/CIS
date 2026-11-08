using Cis.Application.Mco;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Mco;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Refund Listing against the real databases (skipped without CIS_TEST_DB); the credit note is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class McoRefundListingDatabaseTests
{
    private static readonly McoRefundRequest Raya = new()
    {
        Version = 1, DateSelection = 0, From = "01/01/2021", To = "28/07/2021", GroupCompany = "RAYA AIRWAYS SDN BHD", Status = "N",
    };

    private static McoRefundService Service() =>
        new(TestDb.Sessions(), new McoRefundRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()), TimeProvider.System);

    [Fact]
    public async Task Search_checks_listing_and_export()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = Service();
        Assert.Equal(McoRefundService.NoVersion, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new McoRefundRequest()))).Message);
        Assert.Equal(McoRefundService.NoEffDate,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(Raya with { To = "" }))).Message);
        Assert.Equal(McoRefundService.NoCompany,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(Raya with { GroupCompany = "" }))).Message);
        var r = await service.SearchAsync(Raya);
        Assert.Equal(34, r.Principals);
        Assert.Equal(748.95m, r.PrincipalRefund);
        Assert.Equal(r.PrincipalRefund + r.SupplementaryRefund, r.TotalRefund);
        var file = await service.ExportAsync(Raya);
        Assert.EndsWith(".xlsx", file.FileName);
    }

    [Fact]
    public async Task Credit_note_numbers_and_stamps_rolled_back()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new McoRefundRepository();
        await using var s = await TestDb.Sessions().BeginAsync();
        var before = await s.Sql().Connection.QuerySingleAsync<string>("SELECT TOP 1 CLNCRNote FROM HISMaintenance.dbo.CLNMCOCRNoteSeq", transaction: s.Sql().Transaction);
        var note = await repo.NextCreditNoteAsync(s);
        Assert.Equal("C" + (int.Parse(before) + 1).ToString("00000"), note);
        var rows = await repo.PrincipalsAsync(s, new McoRefundFilter(null, "RAYA AIRWAYS SDN BHD", null, false, new DateTime(2021, 1, 1), new DateTime(2021, 7, 28)), "N");
        var member = ((string)rows[0]["MBMNumber"]!).Trim();
        await repo.StampAsync(s, member, note, "DEVADMIN");
        var stamped = await s.Sql().Connection.QuerySingleAsync<string>("SELECT MBMCLNCRNote FROM HISDB.dbo.MBMCLNOthers WHERE MBMNumber = @member",
            new { member }, s.Sql().Transaction);
        Assert.Equal(note, stamped.Trim());
    }
}
