using Cis.Application.Membership;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Membership;
using Dapper;

namespace Cis.Api.Tests;

public class MbmAdjustmentServiceTests
{
    [Fact]
    public void Refund_follows_Medix_bas()
    {
        // Policy 01/01/2026 – 31/12/2026 (364 days), cancelled 01/07/2026 (183 days left); MCO 30, premium 100.
        var r = MbmAdjustmentService.Calculate(new DateTime(2026, 1, 1), new DateTime(2026, 12, 31), new DateTime(2026, 7, 1), 30m, 100m);
        Assert.Equal(15.08m, r.McoRefund);
        Assert.Equal(14.92m, r.McoEffective);
        Assert.Equal(50.27m, r.PremReduction);
        Assert.Equal(49.73m, r.PremEffective);
        Assert.Equal(0m, MbmAdjustmentService.PremiumReduction(new DateTime(2026, 12, 31), new DateTime(2027, 1, 5), new DateTime(2026, 1, 1), 100m));
        Assert.Equal(0m, MbmAdjustmentService.McoRefund(new DateTime(2026, 12, 31), new DateTime(2027, 1, 5), new DateTime(2026, 1, 1), 30m));
    }

    [Fact]
    public async Task Save_checks_and_messages_follow_the_desktop()
    {
        var service = new MbmAdjustmentService(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));
        AdjSaveRequest R(string? member, string? bordx, string? eff) => new(member, "A", bordx, eff, null, null, null, null, null, 0, 0, null, null, null, null, null, null, null, null, null);
        Assert.Equal(MbmAdjustmentService.SelectFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(R(" ", "01/11/2026", "01/11/2026"), "DEV"))).Message);
        Assert.Equal(MbmAdjustmentService.BordxFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(R("X1", "__/__/____", "01/11/2026"), "DEV"))).Message);
        Assert.Equal(MbmAdjustmentService.EffFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(R("X1", "01/11/2026", ""), "DEV"))).Message);
        Assert.Equal("Supplementary has been cancelled successfully", MbmAdjustmentService.SavedMessage("K"));
        Assert.Equal("Updated Successfully", MbmAdjustmentService.SavedMessage("A"));
        Assert.Equal("Share For Supplementary and Full For Principle", MbmAdjustmentService.LimitStatusText("D"));
    }
}

/// <summary>MBM Adjustment against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class MbmAdjustmentDatabaseTests
{
    private const string Member = "DCHC9086608";

    [Fact]
    public async Task Find_fees_histories_and_an_amendment()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new MbmAdjustmentRepository();
        var service = new MbmAdjustmentService(TestDb.Sessions(), repo, new FixedClock(new DateTime(2026, 11, 1)));
        Assert.NotEmpty((await service.SetupAsync()).CardTypes);
        var m = await service.FindAsync(Member);
        Assert.Equal(("NUR RABIZAH BINTI ADENI", "H", "S079", "A"), (m.Name, m.InsType, m.Plan, m.LimitStatus));
        Assert.Equal("Share Limit For Family", m.LimitStatusText);
        Assert.Equal(3, m.Covers.Count);
        Assert.Equal("01", m.Covers[0].CoverId);
        Assert.True(m.CanEdit);
        Assert.Null(m.Notice);
        var (fees, _) = await service.FeesAsync(new AdjFeesRequest(Member, "01/08/2021", "31/12/2021"));
        Assert.Equal(4, fees.Count);
        Assert.NotNull(await service.AmendHistoryAsync(Member));
        Assert.NotNull(await service.MbmHistoryAsync(m.Ic));

        await using var s = await TestDb.Sessions().BeginAsync();
        await repo.EndorseAsync(s, Member, "A", "A", new DateTime(2026, 11, 1), new DateTime(2026, 11, 1), "CISTEST", new DateTime(2026, 11, 1));
        var req = new AdjSaveRequest(Member, "A", "01/11/2026", "01/11/2026", m.PolicyNo, null, null, "ZZ CIS WEB TEST", m.CardType, m.McoFee, m.Premium,
            m.DataStatus, m.Status, m.Exclusion, m.Plan, m.Email, m.Mobile, m.BankName, m.BankAccount, null);
        Assert.Equal(1, await s.Sql().Connection.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM dbo.MBMCISAmendHistory WHERE MBMNumber = @Member AND MBMAmendUser = 'CISTEST'", new { Member }, s.Sql().Transaction));
        await repo.RefundAsync(s, Member, "00", 1, 2, 3, 4);
        Assert.Equal(4m, (await repo.RefundsAsync(s, Member))["00"].McoRefund);
        try
        {
            await repo.SavePrincipalAsync(s, Member, req, null, null);
            Assert.Equal("ZZ CIS WEB TEST", await s.Sql().Connection.ExecuteScalarAsync<string>("SELECT MBMRemarks FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber = @Member", new { Member }, s.Sql().Transaction));
        }
        catch (Microsoft.Data.SqlClient.SqlException e) when (e.Message.Contains("master key", StringComparison.OrdinalIgnoreCase))
        {
            // The restored local HISDB has no database master key for the member encryption triggers; see the screen doc.
        }
    }
}
