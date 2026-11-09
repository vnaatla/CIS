using Cis.Application.Payment;
using Cis.Domain.Errors;
using Cis.Infrastructure.Payment;

namespace Cis.Api.Tests;

/// <summary>Payment Tracking against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class PayTrackingDatabaseTests
{
    [Fact]
    public async Task Clinic_and_member_listings()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PayTrackingService(TestDb.Sessions(), new PayTrackingRepository());
        Assert.Equal(PayTrackingService.NoType, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new PayTrackingRequest()))).Message);
        var clinic = await service.SearchAsync(new PayTrackingRequest { ReportType = 0, Selection = "General Listing", PvFrom = "CP100472166", PvTo = "CP100472166" });
        Assert.Equal(21, clinic.Rows.Count);
        var byPv = await service.SearchAsync(new PayTrackingRequest { ReportType = 0, Selection = "By PV No", PvFrom = "CP100472166", PvTo = "CP100472166" });
        Assert.Equal(clinic.TotalPaid, Assert.Single(byPv.Rows)["CLNPaidAmt"]);
        var member = await service.SearchAsync(new PayTrackingRequest { ReportType = 1, Selection = "By Bordx", PvFrom = "CP100471601", PvTo = "CP100471603" });
        Assert.Equal("Empty", Assert.Single(member.Rows)["BordxRefNo"]);
    }
}
