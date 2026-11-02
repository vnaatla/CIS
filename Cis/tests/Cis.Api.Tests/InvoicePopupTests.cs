using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Invoice;

namespace Cis.Api.Tests;

public class InvoicePopupServiceTests
{
    [Fact]
    public async Task Case_search_needs_three_characters()
    {
        var service = new InvoicePopupService(new FakeSessions(), null!);
        Assert.Equal(InvoicePopupService.TooShort, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.CasesAsync(false, 3, "AB"))).Message);
    }

    [Fact]
    public void Case_rows_follow_ListCases()
    {
        var r = InvoicePopupService.Row(new() { ["CASNumber"] = "K1", ["MBMNumber"] = "M1", ["CoverID"] = "01", ["MBMName"] = "P", ["MBMCName"] = "",
            ["ClinicCode"] = "T301 ", ["CLNName"] = "KLINIK" });
        Assert.Equal((InvoicePopupService.Deleted, "T301 - KLINIK"), (r.Name, r.Clinic));
        Assert.Equal("P", InvoicePopupService.Row(new() { ["CoverID"] = "00", ["MBMName"] = "P " }).Name);
    }
}

/// <summary>Pop-ups against the real databases (skipped without CIS_TEST_DB).</summary>
[Collection(DatabaseCollection.Name)]
public class InvoicePopupDatabaseTests
{
    [Fact]
    public async Task Icd_and_case_searches()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new InvoicePopupService(TestDb.Sessions(), new InvoicePopupRepository());
        Assert.Contains(await service.IcdAsync("305", null, null, "AND"), r => r.ReasonCode == "305");
        var dental = await service.IcdAsync(null, "DENTAL", null, "AND");
        Assert.All(dental, r => Assert.Contains("DENTAL", r.ReasonDesc!, StringComparison.OrdinalIgnoreCase));
        var either = await service.IcdAsync(null, "DENTAL", "OPTICAL", "OR");
        Assert.True(either.Count > dental.Count);
        await Assert.ThrowsAsync<NotFoundException>(() => service.IcdAsync("ZZZZ", null, null, "AND"));

        var cases = await service.CasesAsync(false, 5, "EA1258687");
        Assert.Contains(cases, c => c.CasNumber == "EA12586874" && c.Clinic.StartsWith("T301", StringComparison.Ordinal));
        await Assert.ThrowsAsync<NotFoundException>(() => service.CasesAsync(false, 5, "ZZZZZZ"));
    }
}
