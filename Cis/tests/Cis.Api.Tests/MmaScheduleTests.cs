using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Invoice;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

public class MmaScheduleRulesTests
{
    [Fact]
    public void Filters_and_rows_follow_the_desktop()
    {
        var f = MmaScheduleService.Reduce(new MmaFilters(" brain ", "a01", null, "SCMO5", "X", " b ", "", null, " 7030 ", null));
        Assert.Equal(new MmaFilters("BRAIN", "A01", null, "CMO5", "", "B", null, null, "7030", null), f);
        var b = MmaScheduleService.Row(new() { ["ProcSurgCat"] = "SCMO5", ["ProcPart"] = "B", ["Surgeon13thNew"] = 7030m, ["PartACost"] = 99m });
        Assert.Equal(("CMO5", "7030", ""), (b.SurgCat, b.SurgeonFee, b.PartAFee));
        var a = MmaScheduleService.Row(new() { ["ProcPart"] = "A", ["Surgeon13thNew"] = 7030m, ["PartACost"] = 99m });
        Assert.Equal(("", "99"), (a.SurgeonFee, a.PartAFee));
    }
}

/// <summary>MMA schedule against HISMaintenance (skipped without CIS_TEST_DB).</summary>
[Collection(DatabaseCollection.Name)]
public class MmaScheduleDatabaseTests
{
    [Fact]
    public async Task Search_and_export()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new MmaScheduleService(TestDb.Sessions(), new MmaScheduleRepository(), exporter);
        var filters = new MmaFilters("brain", "A01", null, null, null, null, null, null, null, null);
        var rows = await service.SearchAsync(filters);
        Assert.Contains(rows, r => r.Opcs == "A0110" && r.SurgCat == "CMO5" && r.SurgeonFee == "7030" && r.Part == "B");
        Assert.True((await service.ExportAsync(filters)).Content.Length > 0);
        await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(filters with { Opcs = "ZZZZ" }));
    }
}
