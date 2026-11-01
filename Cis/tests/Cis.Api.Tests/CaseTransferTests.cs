using Cis.Application.Membership;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Membership;
using Dapper;

namespace Cis.Api.Tests;

public class CaseTransferServiceTests
{
    private readonly CaseTransferService _service = new(new FakeSessions(), null!);

    [Fact]
    public async Task Checks_use_the_desktop_messages_in_order()
    {
        Assert.Equal(CaseTransferService.CaseFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.FindAsync(" "))).Message);
        Assert.Equal(CaseTransferService.MemberFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.MemberAsync("", "00"))).Message);
        Assert.Equal(CaseTransferService.CoverFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.MemberAsync("DCGB3518731", ""))).Message);
        Assert.Equal(CaseTransferService.CaseFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            _service.SaveAsync(new CaseTransferRequest("", "DCGB3518731", "00")))).Message);
    }
}

/// <summary>Case Transfer against the real CISDB / HISDB (skipped without CIS_TEST_DB); the transfer is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class CaseTransferDatabaseTests
{
    private const string Case = "K001596118";

    [Fact]
    public async Task A_case_shows_its_member_and_moves_to_another_member()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new CaseTransferRepository();
        await using var s = await TestDb.Sessions().BeginAsync();
        var c = await repo.CaseAsync(s, Case);
        Assert.NotNull(c);
        Assert.Equal(("DCGB3518731", "00"), (c!.MbmNumber, c.CovId));
        Assert.False(string.IsNullOrEmpty(c.Patient));
        Assert.Null(await repo.CaseAsync(s, "ZZNOCASE"));
        Assert.Equal("NUR RABIZAH BINTI ADENI", await repo.MemberNameAsync(s, "DCHC9086608", "00"));
        Assert.False(string.IsNullOrEmpty(await repo.MemberNameAsync(s, "DCHC9086608", "01")));
        Assert.Equal(1, await repo.SaveAsync(s, Case, "DCHC9086608", "01"));
        Assert.Equal("DCHC9086608", await s.Sql().Connection.ExecuteScalarAsync<string>("SELECT MBMnumber FROM dbo.CISCAS WHERE CASNumber = @Case", new { Case }, s.Sql().Transaction));
    }
}
