using Cis.Api.Auth;
using Cis.Api.Auth.Login;
using Cis.Api.Common;
using Cis.Api.Data;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

public class LegacyCryptTests
{
    [Fact]
    public void Shifts_each_character_by_its_position_squared()
    {
        // 'A' (65) + 1, 'A' + 4, 'A' + 9
        Assert.Equal("BEJ", LegacyCrypt.Encrypt("AAA"));
    }

    [Fact]
    public void Wraps_above_255_like_the_desktop()
    {
        // position 16: 'z' (122) + 256 = 378 - 255 = 123 '{'
        Assert.Equal('{', LegacyCrypt.Encrypt(new string('z', 16))[15]);
    }

    [Fact]
    public void Matches_ignores_case_and_padding()
    {
        var stored = LegacyCrypt.Encrypt("secret").ToUpperInvariant() + "  ";
        Assert.True(LegacyCrypt.Matches("SECRET".ToLowerInvariant(), stored));
        Assert.False(LegacyCrypt.Matches("other", stored));
        Assert.False(LegacyCrypt.Matches("secret", null));
    }
}

public class LoginValidatorTests
{
    [Fact]
    public void Login_requires_the_user_name()
    {
        Assert.Equal(LoginMessages.EnterUserName, new LoginRequestValidator().FirstError(new LoginRequest(" ", "x")));
        Assert.Null(new LoginRequestValidator().FirstError(new LoginRequest("A", "")));
    }

    [Theory]
    [InlineData("", "p", "n", "n", LoginMessages.EnterUserName)]
    [InlineData("u", "", "n", "n", LoginMessages.EnterPassword)]
    [InlineData("u", "p", "", "n", LoginMessages.EnterNewPassword)]
    [InlineData("u", "p", "n", "", LoginMessages.EnterConfirmPassword)]
    [InlineData("u", "p", "n1", "n2", LoginMessages.PasswordsDiffer)]
    [InlineData("u", "p", "12345678901", "12345678901", LoginMessages.NewPasswordTooLong)]
    public void Change_password_reports_the_first_problem_in_the_desktop_order(string user, string pwd, string next, string confirm, string message)
    {
        Assert.Equal(message, new ChangePasswordRequestValidator().FirstError(new ChangePasswordRequest(user, pwd, next, confirm)));
    }
}

public class LoginLogFileTests
{
    [Fact]
    public void Writes_the_line_like_vb_write_statement()
    {
        // Line written by the VB6 desktop on 31-10-2026 at 21:10:35 (LocalRun\FileLogin\LoginCIS.txt).
        Assert.Equal("\"Login\",\"DEVADMIN\",#2026-10-31#,\"09:10:35 pm\",\r\n", LoginLogFile.Line("DEVADMIN", new DateTime(2026, 10, 31, 21, 10, 35)));
    }

    [Fact]
    public async Task Appends_to_the_configured_file_and_skips_when_not_configured()
    {
        var path = Path.Combine(Path.GetTempPath(), $"cis-login-{Guid.NewGuid():N}", "LoginCIS.txt");
        var file = new LoginLogFile(Options.Create(new SessionTrackingOptions { LoginLogPath = path }), NullLogger<LoginLogFile>.Instance);
        await file.AppendAsync("A", new DateTime(2026, 1, 2, 3, 4, 5));
        await file.AppendAsync("B", new DateTime(2026, 1, 2, 15, 4, 5));
        Assert.Equal("\"Login\",\"A\",#2026-01-02#,\"03:04:05 am\",\r\n\"Login\",\"B\",#2026-01-02#,\"03:04:05 pm\",\r\n", await File.ReadAllTextAsync(path));
        Directory.Delete(Path.GetDirectoryName(path)!, true);

        await new LoginLogFile(Options.Create(new SessionTrackingOptions()), NullLogger<LoginLogFile>.Instance).AppendAsync("A", DateTime.Now);
    }
}

public class LoginServiceTests
{
    private const string Password = "Gen12345";
    private static readonly DateTime Now = new(2026, 10, 31, 21, 10, 35);

    private sealed class FakeRepository : ILoginRepository
    {
        public Dictionary<string, UserRecord> Users { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> Suspended { get; } = [];
        public List<(string User, string Time, string Date)> Logins { get; } = [];
        public List<string> Logouts { get; } = [];
        public List<(string User, string Password)> Passwords { get; } = [];

        public Task<UserRecord?> GetActiveUserAsync(IDbSession session, string userCode) => Task.FromResult(Users.GetValueOrDefault(userCode));

        public Task SuspendAsync(IDbSession session, string userCode)
        {
            Suspended.Add(userCode);
            return Task.CompletedTask;
        }

        public Task RecordLoginAsync(IDbSession session, string userCode, string loginTime, string loginDate)
        {
            Logins.Add((userCode, loginTime, loginDate));
            return Task.CompletedTask;
        }

        public Task RecordLogoutAsync(IDbSession session, string userCode)
        {
            Logouts.Add(userCode);
            return Task.CompletedTask;
        }

        public Task ChangePasswordAsync(IDbSession session, string userCode, string scrambledPassword)
        {
            Passwords.Add((userCode, scrambledPassword));
            return Task.CompletedTask;
        }
    }

    private sealed class FakeLog : ILoginLog
    {
        public List<string> Lines { get; } = [];

        public Task AppendAsync(string userCode, DateTime now)
        {
            Lines.Add(LoginLogFile.Line(userCode, now));
            return Task.CompletedTask;
        }
    }

    private readonly FakeSessions _sessions = new();
    private readonly FakeRepository _repository = new();
    private readonly FakeLog _log = new();
    private readonly LoginService _service;

    public LoginServiceTests()
    {
        _repository.Users["DEVADMIN"] = new UserRecord("DEVADMIN  ", "Dev Admin", LegacyCrypt.Encrypt(Password).ToUpperInvariant(), "0111");
        _service = new LoginService(_sessions, _repository, new LoginAttempts(new MemoryCache(new MemoryCacheOptions())), _log,
            new TokenService(Options.Create(new AuthOptions { JwtKey = new string('k', 40) })), new FixedClock(Now), NullLogger<LoginService>.Instance);
    }

    [Fact]
    public async Task Success_records_the_login_like_the_desktop_and_returns_the_access_flags()
    {
        var result = await _service.LoginAsync(new LoginRequest("devadmin", Password), "1.2.3.4");

        Assert.Equal(new SessionUser("DEVADMIN", "Dev Admin", "0111"), result.User);
        Assert.False(string.IsNullOrEmpty(result.Token));
        Assert.Equal(("devadmin", "9:10:35 PM", "31/10/2026"), Assert.Single(_repository.Logins));
        Assert.Equal("\"Login\",\"devadmin\",#2026-10-31#,\"09:10:35 pm\",\r\n", Assert.Single(_log.Lines));
        Assert.True(_sessions.Last.Committed);
    }

    [Fact]
    public async Task Unknown_user_is_refused_with_the_desktop_message_and_changes_nothing()
    {
        var e = await Assert.ThrowsAsync<UnauthorizedException>(() => _service.LoginAsync(new LoginRequest("NOBODY", Password), null));
        Assert.Equal(LoginMessages.InvalidUserName, e.Message);
        Assert.Empty(_repository.Logins);
        Assert.False(_sessions.Last.Committed);
    }

    [Fact]
    public async Task Three_wrong_passwords_then_the_next_attempt_suspends_the_account_whatever_the_password()
    {
        for (var i = 0; i < LoginService.MaxFailures; i++)
        {
            var wrong = await Assert.ThrowsAsync<UnauthorizedException>(() => _service.LoginAsync(new LoginRequest("DEVADMIN", "bad"), "pc1"));
            Assert.Equal(LoginMessages.InvalidPassword, wrong.Message);
        }
        Assert.Empty(_repository.Suspended);

        var suspended = await Assert.ThrowsAsync<UnauthorizedException>(() => _service.LoginAsync(new LoginRequest("DEVADMIN", Password), "pc1"));
        Assert.Equal(LoginMessages.Suspended, suspended.Message);
        Assert.Equal("DEVADMIN", Assert.Single(_repository.Suspended));
        Assert.Empty(_repository.Logins);
    }

    [Fact]
    public async Task Failures_are_counted_per_client_and_cleared_by_a_successful_login()
    {
        await Assert.ThrowsAsync<UnauthorizedException>(() => _service.LoginAsync(new LoginRequest("DEVADMIN", "bad"), "pc1"));
        await Assert.ThrowsAsync<UnauthorizedException>(() => _service.LoginAsync(new LoginRequest("DEVADMIN", "bad"), "pc1"));
        await _service.LoginAsync(new LoginRequest("DEVADMIN", Password), "pc1");
        for (var i = 0; i < LoginService.MaxFailures; i++)
            await Assert.ThrowsAsync<UnauthorizedException>(() => _service.LoginAsync(new LoginRequest("DEVADMIN", "bad"), "pc2"));
        await _service.LoginAsync(new LoginRequest("DEVADMIN", Password), "pc1");
        Assert.Empty(_repository.Suspended);
    }

    [Fact]
    public async Task Change_password_stores_the_scrambled_value_so_the_new_password_works()
    {
        var message = await _service.ChangePasswordAsync(new ChangePasswordRequest("DEVADMIN", Password, "New1", "New1"), null);

        Assert.Equal(LoginMessages.PasswordUpdated, message);
        var (user, stored) = Assert.Single(_repository.Passwords);
        Assert.Equal("DEVADMIN  ", user);
        Assert.True(LegacyCrypt.Matches("New1", stored));
        Assert.True(_sessions.Last.Committed);
    }

    [Fact]
    public async Task Change_password_with_a_wrong_current_password_is_refused()
    {
        var e = await Assert.ThrowsAsync<UnauthorizedException>(() =>
            _service.ChangePasswordAsync(new ChangePasswordRequest("DEVADMIN", "wrong", "New1", "New1"), null));
        Assert.Equal(LoginMessages.NoMatch, e.Message);
        Assert.Empty(_repository.Passwords);
    }

    [Fact]
    public async Task Logout_decrements_the_session_count()
    {
        await _service.LogoutAsync("DEVADMIN");
        Assert.Equal("DEVADMIN", Assert.Single(_repository.Logouts));
        Assert.True(_sessions.Last.Committed);
    }
}
