using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Invoice;

namespace Cis.Api.Tests;

public class MemberSuppRulesTests
{
    private static MemberSuppRows Rows(string suppType = "B", string annSupp = "Y", bool coverLimits = true) => new(
        new() { ["MBMName"] = "P", ["HLTCode"] = "H", ["MBMCLNDataStatus"] = "A", ["MBMCLNStatus"] = "A", ["MBMCLNMCDaysBal"] = 3 },
        null,
        new() { ["PLNType"] = "A001", ["MBMCLNBalLmt"] = 500m, ["MBMCLNVisitLmt"] = 40m, ["MBMAccumAmt"] = 10m },
        new() { ["AnnLmtIND"] = "Y", ["VisLmtIND"] = "Y", ["ANNLmtIndSupp"] = annSupp, ["VISLmtIndSupp"] = "N", ["PDentalStatus"] = "P", ["SDentalStatus"] = "S" },
        25m, suppType,
        new() { ["MBMCName"] = "CHILD", ["MBMCRELCode"] = "C" },
        coverLimits ? [new() { ["MBMCCoverID"] = "01", ["MBMCCLNBalLmt"] = 300m, ["MBMCCLNVisitLmt"] = 30m, ["MBMCCLNStatus"] = "A", ["MBMCAccumAmt"] = 7m }] : []);

    [Fact]
    public void Principal_and_cover_limits_follow_MemberSuppInfo()
    {
        var p = MemberSuppService.Build("M1", "00", Rows());
        Assert.Equal(("500", "40", "P", "25", "Y", 3m), (p.BalanceLimit, p.VisitLimit, p.DentalStatus, p.CoPay, p.AnnualIndicator, p.McDaysBalance));
        var d = MemberSuppService.Build("M1", "00", Rows("D"));
        Assert.Equal(("FULL", "N"), (d.BalanceLimit, d.AnnualIndicator));
        // Shared limit type B: cover 02 uses cover 01's limits; VISLmtIndSupp N → visit FULL.
        var c = MemberSuppService.Build("M1", "02", Rows());
        Assert.Equal(("CHILD", "300", "FULL", 7m, "S"), (c.PatientName, c.BalanceLimit, c.VisitLimit, c.AccumAmount, c.DentalStatus));
        Assert.Equal("D", MemberSuppService.Build("M1", "02", Rows(coverLimits: false)).DataStatus);
        // ANNLmtIndSupp N: the supplementary type stays A (no AnnualLimit lookup) → the principal's limits.
        Assert.Equal(("A", "500"), (MemberSuppService.Build("M1", "02", Rows(annSupp: "N")).SuppLimitType, MemberSuppService.Build("M1", "02", Rows(annSupp: "N")).BalanceLimit));
        Assert.Equal(MemberSuppService.NoCover, Assert.Throws<NotFoundException>(() => MemberSuppService.Build("M1", "02", Rows() with { Cover = null })).Message);
        Assert.Equal(MemberSuppService.NoPlan, Assert.Throws<NotFoundException>(() => MemberSuppService.Build("M1", "00", Rows() with { Plan = null })).Message);
    }
}

/// <summary>MemberSuppInfo against HISDB / CISDB (skipped without CIS_TEST_DB).</summary>
[Collection(DatabaseCollection.Name)]
public class MemberSuppDatabaseTests
{
    [Fact]
    public async Task A_real_member_and_cover()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new MemberSuppService(TestDb.Sessions(), new MemberSuppRepository());
        var p = await service.InfoAsync("dchc9086608", "00");
        Assert.Equal("NUR RABIZAH BINTI ADENI", p.PatientName);
        Assert.False(string.IsNullOrEmpty(p.PlanCode));
        var c = await service.InfoAsync("DCHC9086608", "01");
        Assert.NotEqual(p.PatientName, c.PatientName);
        await Assert.ThrowsAsync<NotFoundException>(() => service.InfoAsync("ZZNOMEMBER", "00"));
    }
}
