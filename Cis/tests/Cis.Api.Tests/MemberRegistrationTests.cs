using Cis.Application.Membership;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Membership;
using Microsoft.Data.SqlClient;

namespace Cis.Api.Tests;

public class MemberRegistrationServiceTests
{
    [Theory]
    [InlineData("2026-01-01", "1982-06-15", "43Y")]
    [InlineData("2026-07-01", "1982-06-15", "43Y")] // birth day 15 > 1: one year less even though June has passed (desktop quirk)
    [InlineData("2026-07-20", "1982-06-15", "44Y")]
    [InlineData("2026-07-01", "1982-06-30", "43Y")]
    [InlineData("2026-07-01", "2015-03-01", "0Y")]
    [InlineData("2026-01-01", "2015-03-01", "-1Y")]
    [InlineData("2026-01-01", "2026-03-01", "-1D")]
    public void GetAge_follows_Medix_bas(string eff, string dob, string age) =>
        Assert.Equal(age, MemberRegistrationService.GetAge(DateTime.Parse(eff, System.Globalization.CultureInfo.InvariantCulture),
            DateTime.Parse(dob, System.Globalization.CultureInfo.InvariantCulture)));

    [Fact]
    public async Task Register_checks_run_in_the_desktop_order()
    {
        var service = new MemberRegistrationService(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));
        var ok = new RegistrationRequest("820615136152", "ZZ TEST", "M", "F", "Crew", "E1", "01/11/2026", "31/12/2026", "15/06/1982", "SACOFA SDN BHD", "IT");
        Assert.Equal("Please enter IC No! ", (await Assert.ThrowsAsync<BusinessRuleException>(() => service.RegisterAsync(ok with { Ic = "" }, "DEV"))).Message);
        Assert.Equal("Please enter Staff Category! ", (await Assert.ThrowsAsync<BusinessRuleException>(() => service.RegisterAsync(ok with { StaffCategory = " " }, "DEV"))).Message);
        Assert.Equal("Please enter Birth Date! ", (await Assert.ThrowsAsync<BusinessRuleException>(() => service.RegisterAsync(ok with { Dob = "__/__/____" }, "DEV"))).Message);
        Assert.Equal("Please enter Department! ", (await Assert.ThrowsAsync<BusinessRuleException>(() => service.RegisterAsync(ok with { Department = null }, "DEV"))).Message);
    }
}

/// <summary>Registration look-ups and the save against the real databases (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class MemberRegistrationDatabaseTests
{
    [Fact]
    public async Task Lists_limits_age_band_member_by_ic_and_register()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new MemberRegistrationRepository();
        var service = new MemberRegistrationService(TestDb.Sessions(), repo, new FixedClock(new DateTime(2026, 11, 1)));
        var setup = await service.SetupAsync();
        Assert.Contains(setup.Races, r => r.Code == "M");
        Assert.Contains(setup.Categories, c => c.Category == "Crew" && c.PlanCode == "G007");
        Assert.NotEmpty(setup.GroupCompanies);
        Assert.NotNull(await service.LimitsAsync("Crew", "01/11/2026"));
        var adult = await service.AgeAsync("01/11/2026", "15/06/1982");
        Assert.Equal(("43Y", "A"), (adult.Age, adult.AdultChild));
        Assert.Equal("C", (await service.AgeAsync("01/11/2026", "15/06/2020")).AdultChild);
        var found = await service.FindByIcAsync("820615136152");
        Assert.False(found.Active);
        Assert.Equal("NUR RABIZAH BINTI ADENI", found.Name);
        Assert.NotEmpty(await service.DepartmentsAsync("A"));

        await using var s = await TestDb.Sessions().BeginAsync();
        var member = new NewMember("ZZ0000000001", "ZZ CIS WEB TEST", "M", "A", "02", new DateTime(1982, 6, 15), "M", "ZZ GROUP", "ZZ1", "ZZ DEPT",
            "G007", "G", new DateTime(2026, 11, 1), new DateTime(2026, 12, 31), 0, 0);
        try
        {
            Assert.Matches("^DCG[A-Z][0-9]{5}[0-9]{2}$", await repo.RegisterAsync(s, member, "CISTEST", new DateTime(2026, 11, 1)));
        }
        catch (SqlException e) when (e.Message.Contains("master key", StringComparison.OrdinalIgnoreCase))
        {
            // The restored local HISDB has no database master key for the MBM IC-encryption trigger (trgICEncrypt); see the screen doc.
        }
    }
}
