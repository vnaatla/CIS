using Cis.Application.Abstractions;
using Cis.Application.Membership;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Membership;

namespace Cis.Api.Tests;

public class MemberSearchServiceTests
{
    private sealed class FakeRepository : IMemberSearchRepository
    {
        public List<MemberSearchRow> Principals { get; } = [];
        public List<MemberSearchRow> Supps { get; } = [];
        public int SuppCalls { get; private set; }

        public Task<List<MemberSearchRow>> PrincipalsAsync(IDbSession session, MemberSearchBy by, string text, int max) => Task.FromResult(Principals.ToList());

        public Task<List<MemberSearchRow>> SuppsAsync(IDbSession session, MemberSearchBy by, string text, int max)
        {
            SuppCalls++;
            return Task.FromResult(Supps.ToList());
        }

        public Task<List<MemberSearchRow>> ByPrefixAsync(IDbSession session, string prefix) => Task.FromResult(new List<MemberSearchRow>());
    }

    private static MemberSearchRow Row(string no, bool supp) => new(no, supp ? "01" : "00", "NAME", "IC", null, "A", "A", null, null, "GRP", "DEPT", supp);

    private readonly FakeRepository _repository = new();
    private MemberSearchService Service => new(new FakeSessions(), _repository);

    [Fact]
    public async Task Short_text_and_no_match_give_the_desktop_messages()
    {
        Assert.Equal(MemberSearchService.TooShort, (await Assert.ThrowsAsync<BusinessRuleException>(() => Service.SearchAsync(MemberSearchBy.Name, " ab "))).Message);
        Assert.Equal(MemberSearchService.NotFound, (await Assert.ThrowsAsync<NotFoundException>(() => Service.SearchAsync(MemberSearchBy.Name, "ZZZZ"))).Message);
        Assert.Equal(MemberSearchService.NotFound05, (await Assert.ThrowsAsync<NotFoundException>(() => Service.ByPrefixAsync("ZZ"))).Message);
    }

    [Fact]
    public async Task Principals_come_first_with_the_desktop_columns_and_supps_only_for_number_name_and_ic()
    {
        _repository.Principals.Add(Row("P1", false));
        _repository.Supps.Add(Row("P1", true));
        var result = await Service.SearchAsync(MemberSearchBy.Name, "LIM");
        Assert.Equal((1, 1, false), (result.Principals, result.Supps, result.Capped));
        Assert.Equal([false, true], result.Rows.Select(r => r.Supp));
        Assert.Null(result.Rows[0].PolicyStatus);
        Assert.Null(result.Rows[0].GroupCompany);
        Assert.Equal("A", result.Rows[1].PolicyStatus);

        await Service.SearchAsync(MemberSearchBy.EmployeeNo, "E123");
        Assert.Equal(1, _repository.SuppCalls);
    }
}

/// <summary>The pop-up procedures against the real HISDB / HISMaintenance (skipped without CIS_TEST_DB).</summary>
[Collection(DatabaseCollection.Name)]
public class MemberSearchDatabaseTests
{
    [Fact]
    public async Task A_member_is_found_by_number_ic_and_prefix()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().OpenAsync();
        var any = await Dapper.SqlMapper.ExecuteScalarAsync<string>(s.Sql().Connection,
            "SELECT TOP 1 c.MBMNumber FROM cis.Hdb_MBMCLNOthers c JOIN cis.Maint_MBMCrossReference x ON x.MBMNumber = c.MBMNumber WHERE LEN(x.MBMIcBcPp) >= 12");
        Assert.NotNull(any);
        var repo = new MemberSearchRepository();
        var byNumber = Assert.Single(await repo.PrincipalsAsync(s, MemberSearchBy.MemberNo, any!, 10));
        Assert.Contains(await repo.PrincipalsAsync(s, MemberSearchBy.IC, byNumber.Ic!, 50), r => r.MbmNumber == any);
        Assert.Contains(await repo.ByPrefixAsync(s, any![..Math.Min(any.Length, 8)]), r => r.MbmNumber == any);
    }
}
