using Cis.Application.Bordx;
using Cis.Domain.Errors;
using Cis.Infrastructure.Bordx;

namespace Cis.Api.Tests;

/// <summary>BSN-PrintSummary against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class BordxBsnDatabaseTests
{
    [Fact]
    public async Task Summary_matches_the_desktop_procedure()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BordxBsnService(TestDb.Sessions(), new BordxBsnRepository());
        Assert.Equal(BordxBsnService.NoBordx, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SummaryAsync([" "]))).Message);
        var rows = await service.SummaryAsync(["KDC55481", "KIV03230"]);
        Assert.Equal(2, rows.Count);
        Assert.Equal(772, Convert.ToInt32(rows[0]["NOOFCLAIMS"]));
        Assert.Equal(86652.82m, (decimal)rows[0]["CLMBordxAmt"]!);
    }
}
