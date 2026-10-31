using System.Reflection;
using System.Text.RegularExpressions;
using Cis.Api.Common;
using FluentValidation;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>The layering of docs/architecture.md, enforced (same rules as GENESIS).</summary>
public class ArchitectureTests(ApiFactory factory) : IClassFixture<ApiFactory>
{
    private static readonly Assembly Api = typeof(Program).Assembly;

    private static IEnumerable<Type> Controllers => Api.GetTypes().Where(t => typeof(ControllerBase).IsAssignableFrom(t) && !t.IsAbstract);

    [Fact]
    public void Controllers_hold_no_data_access_and_no_try_catch()
    {
        var banned = new Regex(@"\b(Dapper|SqlConnection|SqlCommand|DbConnections|IDbSessions?|try|catch)\b");
        foreach (var file in Directory.GetFiles(Path.Combine(RepoPaths.CisRoot, "src", "Cis.Api", "Controllers"), "*.cs"))
            Assert.False(banned.IsMatch(File.ReadAllText(file)), $"{Path.GetFileName(file)} touches data access or catches exceptions");
    }

    [Fact]
    public void Controllers_derive_from_the_base_and_use_api_routes()
    {
        foreach (var c in Controllers)
        {
            Assert.True(typeof(ApiControllerBase).IsAssignableFrom(c), $"{c.Name} must derive from ApiControllerBase");
            Assert.StartsWith("api/", c.GetCustomAttribute<RouteAttribute>()?.Template ?? "", StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Only_login_and_change_password_are_anonymous()
    {
        var anonymous = Controllers.SelectMany(c => c.GetMethods().Where(m => m.GetCustomAttribute<AllowAnonymousAttribute>() is not null).Select(m => $"{c.Name}.{m.Name}"))
            .Concat(Controllers.Where(c => c.GetCustomAttribute<AllowAnonymousAttribute>() is not null).Select(c => c.Name));
        Assert.Equal(["AuthController.ChangePassword", "AuthController.Login"], anonymous.Order());
    }

    [Fact]
    public void Services_depend_on_abstractions_only()
    {
        Type[] allowedConcrete = [typeof(Cis.Api.Auth.TokenService), typeof(TimeProvider), typeof(Cis.Api.Data.DbConnections),
            typeof(Cis.Api.Documents.ExcelReportWriter), typeof(Cis.Api.Documents.PdfReportWriter)];
        foreach (var service in Api.GetTypes().Where(t => t.IsClass && !t.IsAbstract && t.Name.EndsWith("Service", StringComparison.Ordinal) && t != typeof(Cis.Api.Auth.TokenService)))
            foreach (var p in service.GetConstructors().SelectMany(c => c.GetParameters()))
            {
                var type = p.ParameterType;
                var ok = type.IsInterface || allowedConcrete.Contains(type) ||
                    (type.IsGenericType && (type.GetGenericTypeDefinition() == typeof(ILogger<>) || type.GetGenericTypeDefinition() == typeof(IOptions<>)));
                Assert.True(ok, $"{service.Name} depends on concrete {type.Name}");
            }
    }

    [Fact]
    public void Every_repository_and_service_interface_is_registered()
    {
        using var scope = factory.Services.CreateScope();
        foreach (var contract in Api.GetTypes().Where(t => t.IsInterface && Regex.IsMatch(t.Name, "^I[A-Z].*(Repository|Service)$")))
            Assert.True(scope.ServiceProvider.GetService(contract) is not null, $"{contract.Name} is not registered");
    }

    [Fact]
    public void Every_request_model_has_a_validator()
    {
        using var scope = factory.Services.CreateScope();
        var requests = Controllers.SelectMany(c => c.GetMethods()).SelectMany(m => m.GetParameters())
            .Select(p => p.ParameterType).Where(t => t.Name.EndsWith("Request", StringComparison.Ordinal)).Distinct();
        foreach (var request in requests)
            Assert.True(scope.ServiceProvider.GetService(typeof(IValidator<>).MakeGenericType(request)) is not null, $"{request.Name} has no validator");
    }

    [Fact]
    public void Every_screen_endpoint_outside_auth_audit_and_lookups_carries_an_access_check()
    {
        string[] shared = ["AuthController", "AuditController", "LookupsController"];
        foreach (var c in Controllers.Where(c => !shared.Contains(c.Name)))
        {
            var classLevel = c.GetCustomAttributes<Cis.Api.Auth.RequireAccessAttribute>().Any();
            foreach (var action in c.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
                Assert.True(classLevel || action.GetCustomAttributes<Cis.Api.Auth.RequireAccessAttribute>().Any(), $"{c.Name}.{action.Name} has no [RequireAccess]");
        }
    }
}
