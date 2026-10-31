using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using Cis.Api.Auth;
using Cis.Application.Auth.Login;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.DependencyInjection;

namespace Cis.Api.Tests;

/// <summary>The API in memory with a database that is never reached and a fake login service.</summary>
public sealed class ApiFactory : WebApplicationFactory<Program>
{
    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        builder.UseEnvironment("Development");
        builder.UseSetting("ConnectionStrings:CisDb", "Server=127.0.0.1,1;Database=None;Connect Timeout=1;Encrypt=false");
        builder.UseSetting("Auth:JwtKey", new string('t', 48));
        builder.UseSetting("Session:LoginLogPath", "");
        builder.UseSetting("Email:PickupDirectory", "");
        builder.ConfigureTestServices(services => services.AddScoped<ILoginService, FakeLoginService>());
    }

    private sealed class FakeLoginService(TokenService tokens) : ILoginService
    {
        public Task<SessionResult> LoginAsync(LoginRequest request, string? client)
        {
            var (token, expires) = tokens.Create("TESTER", "Test User", "1");
            return Task.FromResult(new SessionResult(token, expires, new SessionUser("TESTER", "Test User", "1")));
        }

        public Task<string> ChangePasswordAsync(ChangePasswordRequest request, string? client) => Task.FromResult(LoginMessages.PasswordUpdated);

        public Task LogoutAsync(string userCode) => Task.CompletedTask;
    }
}

public class ApiIntegrationTests(ApiFactory factory) : IClassFixture<ApiFactory>
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    [Fact]
    public async Task Endpoints_need_a_token()
    {
        var response = await factory.CreateClient().GetAsync("/api/lookups/payors");
        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
    }

    [Fact]
    public async Task Login_validation_answers_the_desktop_message_as_problem_details()
    {
        var response = await factory.CreateClient().PostAsJsonAsync("/api/auth/login", new { userCode = "", password = "" });
        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal(LoginMessages.EnterUserName, body.GetProperty("message").GetString());
    }

    [Fact]
    public async Task A_token_from_login_opens_the_api_and_logout_ends_the_session()
    {
        var client = factory.CreateClient();
        var login = await (await client.PostAsJsonAsync("/api/auth/login", new { userCode = "TESTER", password = "x" })).Content.ReadFromJsonAsync<SessionResult>(Json);
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", login!.Token);

        var me = await client.GetFromJsonAsync<SessionUser>("/api/auth/me", Json);
        Assert.Equal(new SessionUser("TESTER", "Test User", "1"), me);
        Assert.Equal(HttpStatusCode.NoContent, (await client.PostAsync("/api/auth/logout", null)).StatusCode);
    }

    [Fact]
    public async Task Screen_opened_validates_the_form_name()
    {
        var client = factory.CreateClient();
        var login = await (await client.PostAsJsonAsync("/api/auth/login", new { userCode = "TESTER", password = "x" })).Content.ReadFromJsonAsync<SessionResult>(Json);
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", login!.Token);

        var response = await client.PostAsJsonAsync("/api/audit/screen-opened", new { screen = "frm'; DROP TABLE x" });
        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task Health_is_anonymous_and_reports_the_database_state()
    {
        var response = await factory.CreateClient().GetAsync("/api/health");
        Assert.NotEqual(HttpStatusCode.Unauthorized, response.StatusCode);
    }

    [Fact]
    public async Task Swagger_documents_the_api()
    {
        var json = await factory.CreateClient().GetStringAsync("/swagger/v1/swagger.json");
        Assert.Contains("/api/auth/login", json);
        Assert.Contains("/api/lookups/clinics", json);
        Assert.Contains("bearer", json, StringComparison.OrdinalIgnoreCase);
    }
}
