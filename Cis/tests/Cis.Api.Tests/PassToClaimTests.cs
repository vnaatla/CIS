using Cis.Application.Claim;
using Cis.Domain.Errors;
using Cis.Infrastructure.Claim;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

public class PassToClaimRulesTests
{
    [Fact]
    public void Rows_and_check_rules_follow_the_desktop()
    {
        Assert.Null(PassToClaimService.Row(new() { ["CoverID"] = "00", ["MBMCoverID"] = "01" }));
        var r = PassToClaimService.Row(new() { ["CoverID"] = "01", ["MBMCoverID"] = "01", ["ClaimStatus"] = "P", ["NoofDay"] = 0, ["CASNumber"] = "K1" })!;
        Assert.Equal(("Passed", "1"), (r.Passed, r.NoOfDay));
        Assert.Equal(PassToClaimService.NoRegister, PassToClaimService.CheckMessage(r));
        var ok = r with { RegisterDate = DateTime.Today, VisitDate = DateTime.Today, ClinicName = "KLINIK" };
        Assert.Null(PassToClaimService.CheckMessage(ok));
        Assert.Equal(PassToClaimService.SentOut, PassToClaimService.CheckMessage(ok with { ClaimStatus = "1" }));
        Assert.Equal(PassToClaimService.NoClinic, PassToClaimService.CheckMessage(ok with { ClinicName = "" }));
    }

    [Fact]
    public async Task Scan_mode_needs_a_case_number()
    {
        var service = new PassToClaimService(new FakeSessions(), null!, null!, new FixedClock(new DateTime(2026, 11, 2)));
        var f = new PassFilters(true, " ", null, null, null, null, null, false, false, false, null, null, null, null, null, null);
        Assert.Equal(PassToClaimService.CaseFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(f))).Message);
    }
}

/// <summary>Pass to Claim against CISWorksheet (skipped without CIS_TEST_DB); the pass is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PassToClaimDatabaseTests
{
    [Fact]
    public async Task Scan_list_and_pass()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new PassToClaimRepository();
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new PassToClaimService(TestDb.Sessions(), repo, exporter, new FixedClock(new DateTime(2026, 11, 2)));
        Assert.NotNull(await service.UsersAsync());
        var scan = new PassFilters(true, "ea12586874", null, "2", null, null, null, false, false, false, null, null, null, null, null, null);
        var rows = await service.SearchAsync(scan);
        var row = Assert.Single(rows);
        Assert.Equal(("EA12586874", "Not", "KLINIK TAN (TAMAN SETIA INDAH)"), (row.CasNumber, row.Passed, row.ClinicName));
        Assert.True((await service.ExportAsync(scan)).Content.Length > 0);
        await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(scan with { CaseNo = "ZZNOCASE" }));

        await using var s = await TestDb.Sessions().BeginAsync();
        await repo.PassAsync(s, "EA12586874", "DEV", new DateTime(2026, 11, 2));
        var passed = await s.Sql().Connection.QuerySingleAsync<(string Status, string By)>(
            "SELECT TOP 1 ClaimStatus, ClaimPassBy FROM CISWorksheet.dbo.ClaimFileTransfer WHERE CASNumber = 'EA12586874' ORDER BY ClaimPassDate DESC",
            transaction: s.Sql().Transaction);
        Assert.Equal(("P", "DEV"), passed);
    }
}
