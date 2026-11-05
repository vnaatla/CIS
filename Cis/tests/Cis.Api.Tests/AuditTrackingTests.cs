using Cis.Application.Audit;
using Cis.Domain.Errors;
using Cis.Infrastructure.Audit;
using Cis.Infrastructure.Documents;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Audit Tracking against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class AuditTrackingDatabaseTests
{
    private static AuditTrackingService Service() =>
        new(TestDb.Sessions(), new AuditTrackingRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()), TimeProvider.System);

    [Fact]
    public async Task Search_lists_one_row_per_medication_and_prints()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = Service();
        var filter = new AuditTrackingRequest { CaseFrom = "EA12586874", CaseTo = "EA12586874" };
        var rows = await service.SearchAsync(filter);
        Assert.Equal(6, rows.Count);
        Assert.True(rows[0].First);
        Assert.Equal("EA12586874", rows[0].Row["CASNumber"]);
        Assert.Equal("LORATADINE 10MG TAB", rows[0].Row["MedInjType"]);
        Assert.Equal("DIABETES", rows[0].Row["ReasonDesc1"]);
        Assert.False(rows[1].First);
        Assert.Null(rows[1].Row["CASNumber"]);
        Assert.Equal("GALVUS MET 50/850 TAB", rows[1].Row["MedInjType"]);

        var none = await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(filter with { PayTo = "M" }));
        Assert.Equal(AuditTrackingService.NoRecord, none.Message);
        var omni = await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(filter with { OmniRefFrom = "A" }));
        Assert.Equal(AuditTrackingService.OmniRefMissing, omni.Message);

        var file = await service.PrintAsync(filter);
        Assert.EndsWith(".xlsx", file.FileName);
        Assert.True(file.Content.Length > 1000);

        var (subsidiaries, departments) = await service.LookupsAsync("BANK SIMPANAN NASIONAL");
        Assert.Contains("JABATAN OPERASI PERBANKAN", departments);
        Assert.NotNull(subsidiaries);
    }
}
