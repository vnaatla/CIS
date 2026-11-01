using Cis.Application.Membership;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Membership;
using Dapper;

namespace Cis.Api.Tests;

public class MbmTopupServiceTests
{
    [Theory]
    [InlineData("DCHC9086608", "DCHC9086608", "00")]
    [InlineData("dchc9086608-01", "DCHC9086608", "01")]
    [InlineData("AXHB1101693*01- 00", "AXHB1101693*01", "00")]
    public void Member_box_splits_into_number_and_cover(string text, string number, string cov) =>
        Assert.Equal((number, cov), MbmTopupService.Split(text));

    [Fact]
    public void Balance_status_and_nothing_to_adjust_follow_the_desktop()
    {
        var rows = new List<Dictionary<string, object?>>
        {
            new() { ["CovId"] = "00", ["GPBalance"] = "3129.00", ["SPBalance"] = "NA" },
        };
        Assert.Equal(3129m, MbmTopupService.Balance(rows, "00", true));
        Assert.Equal(0m, MbmTopupService.Balance(rows, "00", false));
        Assert.Null(MbmTopupService.Balance(rows, "05", true));
        Assert.Equal("C - Cancelled", MbmTopupService.StatusLabel("C"));
    }

    [Fact]
    public async Task Save_refuses_no_change_and_an_empty_member()
    {
        var service = new MbmTopupService(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));
        Assert.Equal(MbmTopupService.Nothing, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SaveAsync(new TopupSaveRequest("DCHC9086608", true, "820615136152", "820615136152", "3129", "", "A"), "DEV"))).Message);
        Assert.Equal(MbmTopupService.EnterMember, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SaveAsync(new TopupSaveRequest(" ", true, "", "", "", "10", "A"), "DEV"))).Message);
    }
}

/// <summary>MBM Topup against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class MbmTopupDatabaseTests
{
    private const string Member = "DCHC9086608";

    [Fact]
    public async Task Find_shows_the_member_and_balance_and_a_topup_writes_history_and_limit()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new MbmTopupRepository();
        var service = new MbmTopupService(TestDb.Sessions(), repo, new FixedClock(new DateTime(2026, 11, 1)));
        var m = await service.FindAsync(Member, true);
        Assert.Equal(("00", "S079"), (m.CovId, m.Plan));
        Assert.False(string.IsNullOrEmpty(m.Name));
        Assert.Equal(3129m, m.BalanceLimit);
        Assert.Matches("^ ?[A-Z0-9*]+- 00$", await service.StepAsync(Member, true));

        await using var s = await TestDb.Sessions().BeginAsync();
        var before = await s.Sql().Connection.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM dbo.MBMCISTopupHistory WHERE MBMNumber = @Member", new { Member }, s.Sql().Transaction);
        await repo.SaveAsync(s, new TopupSave(Member, "00", true, m.IcNo, m.IcNo, 3129m, 100m, 3229m, m.LimitCode), "CISTEST", new DateTime(2026, 11, 1));
        Assert.Equal(before + 1, await s.Sql().Connection.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM dbo.MBMCISTopupHistory WHERE MBMNumber = @Member", new { Member }, s.Sql().Transaction));
        Assert.Equal(3229m, await s.Sql().Connection.ExecuteScalarAsync<decimal>("SELECT MBMCLNBalLmt FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber = @Member", new { Member }, s.Sql().Transaction));
    }
}
