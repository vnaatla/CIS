using Cis.Api.Audit;
using Cis.Api.Auth;
using Cis.Api.Auth.Login;
using Cis.Api.Data;
using Cis.Api.Lookups;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>
/// Real procedures against the local databases, each in a transaction that is rolled back (nothing is kept). Skipped when CIS_TEST_DB is
/// not set; see AGENTS.md for the connection string.
/// </summary>
public class DatabaseTests
{
    private const string DevUser = "DEVADMIN";

    private static async Task<IDbSession?> BeginAsync()
    {
        if (TestDb.Cis is null) return null;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        return await TestDb.Sessions().BeginAsync();
    }

    private static Task<T> Scalar<T>(IDbSession s, string sql, object? p = null) => s.Connection.ExecuteScalarAsync<T>(sql, p, s.Transaction)!;

    [Fact]
    public async Task Login_procedures_find_the_dev_user_and_track_the_session_like_the_desktop()
    {
        await using var s = await BeginAsync();
        if (s is null) return;
        var repo = new LoginRepository();

        var user = await repo.GetActiveUserAsync(s, DevUser);
        Assert.NotNull(user);
        Assert.True(LegacyCrypt.Matches("Gen12345", user!.UsrPassword), "local dev password is Gen12345 (AGENTS.md)");
        Assert.Null(await repo.GetActiveUserAsync(s, "NO-SUCH-USER"));

        var count = await Scalar<int?>(s, "SELECT CIScnt FROM cis.Maint_USR WHERE USRCode = @DevUser", new { DevUser }) ?? 0;
        var tracks = await Scalar<int>(s, "SELECT COUNT(*) FROM cis.Maint_USRTrack WHERE USRCode = @DevUser AND SYSName = 'CIS'", new { DevUser });

        await repo.RecordLoginAsync(s, DevUser, "9:10:35 PM", "31/10/2026");
        Assert.Equal(count + 1, await Scalar<int>(s, "SELECT CIScnt FROM cis.Maint_USR WHERE USRCode = @DevUser", new { DevUser }));
        Assert.Equal(tracks + 1, await Scalar<int>(s, "SELECT COUNT(*) FROM cis.Maint_USRTrack WHERE USRCode = @DevUser AND SYSName = 'CIS'", new { DevUser }));
        Assert.Equal("31/10/2026", await Scalar<string>(s, "SELECT TOP 1 LoginDate FROM cis.Maint_USRTrack WHERE USRCode = @DevUser ORDER BY TrackId DESC", new { DevUser }));

        await repo.RecordLogoutAsync(s, DevUser);
        Assert.Equal(count, await Scalar<int>(s, "SELECT CIScnt FROM cis.Maint_USR WHERE USRCode = @DevUser", new { DevUser }));
    }

    [Fact]
    public async Task Suspend_and_change_password_update_usr()
    {
        await using var s = await BeginAsync();
        if (s is null) return;
        var repo = new LoginRepository();

        await repo.ChangePasswordAsync(s, DevUser, LegacyCrypt.Encrypt("Tmp1").ToUpperInvariant());
        Assert.True(LegacyCrypt.Matches("Tmp1", (await repo.GetActiveUserAsync(s, DevUser))!.UsrPassword));

        await repo.SuspendAsync(s, DevUser);
        Assert.Equal("S", await Scalar<string>(s, "SELECT USRStatus FROM cis.Maint_USR WHERE USRCode = @DevUser", new { DevUser }));
        Assert.Null(await repo.GetActiveUserAsync(s, DevUser));
    }

    [Fact]
    public async Task Action_and_error_logs_write_the_legacy_tables()
    {
        await using var s = await BeginAsync();
        if (s is null) return;

        await new ActionLogRepository().InsertAsync(s, DevUser, ActionLogService.ScreenOpened, "frmCisWebTest");
        Assert.Equal(1, await Scalar<int>(s, "SELECT COUNT(*) FROM cis.Hdb_ActionLog WHERE Notes = 'frmCisWebTest' AND ProjType = 'CIS' AND ActionId = 12"));

        await s.ExecAsync("cis.ErrorLog_Write", new { formName = "/api/test", methodName = "GET", description = "cis web test" });
        Assert.Equal(1, await Scalar<int>(s, "SELECT COUNT(*) FROM cis.Maint_ErrorLog WHERE SysName = 'CIS' AND ErrDesc = 'cis web test'"));
    }

    [Fact]
    public async Task Lookups_return_the_desktop_lists()
    {
        await using var s = await BeginAsync();
        if (s is null) return;
        var repo = new LookupRepository();

        var payors = await repo.PayorsAsync(s, null);
        Assert.NotEmpty(payors);
        Assert.All(payors, p => Assert.False(string.IsNullOrEmpty(p.Code)));
        Assert.NotEmpty(await repo.HealthTypesAsync(s));
        Assert.NotEmpty(await repo.InsuredTypesAsync(s));
        Assert.NotEmpty(await repo.GroupCompaniesAsync(s, null));

        var panel = await repo.ClinicsAsync(s, true, null, 20);
        Assert.InRange(panel.Count, 1, 20);
        var match = await repo.ClinicsAsync(s, false, panel[0].Code, 5);
        Assert.Contains(match, c => c.Code == panel[0].Code);
    }
}
