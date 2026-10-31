using System.Reflection;
using System.Text.RegularExpressions;
using Cis.Api.Auth;
using Cis.Api.Common;
using FluentValidation;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>
/// The Clean Architecture boundaries of docs/architecture.md. The compiler already stops a project from using one it does not reference;
/// these tests stop someone from adding the wrong reference, and check what the compiler cannot (thin controllers, registrations, validators).
/// </summary>
public class ArchitectureTests(ApiFactory factory) : IClassFixture<ApiFactory>
{
    private static readonly Assembly Domain = typeof(Cis.Domain.Security.AccessRight).Assembly;
    private static readonly Assembly Application = typeof(Cis.Application.DependencyInjection).Assembly;
    private static readonly Assembly Infrastructure = typeof(Cis.Infrastructure.DependencyInjection).Assembly;
    private static readonly Assembly Api = typeof(Program).Assembly;

    private static IEnumerable<Type> Controllers => Api.GetTypes().Where(t => typeof(ControllerBase).IsAssignableFrom(t) && !t.IsAbstract);

    private static string[] References(Assembly assembly) => assembly.GetReferencedAssemblies().Select(a => a.Name!).ToArray();

    [Fact]
    public void Domain_references_no_other_project_and_no_library()
    {
        Assert.All(References(Domain), name => Assert.True(name is "System.Runtime" || name.StartsWith("System.", StringComparison.Ordinal),
            $"Cis.Domain must only use the BCL, but references {name}"));
    }

    [Fact]
    public void Application_references_only_the_domain_and_abstractions()
    {
        var refs = References(Application);
        Assert.Contains("Cis.Domain", refs);
        string[] banned = ["Cis.Infrastructure", "Cis.Api", "Microsoft.Data.SqlClient", "Dapper", "ClosedXML", "PdfSharp", "MigraDoc", "MailKit", "MimeKit"];
        Assert.All(refs, name =>
        {
            Assert.DoesNotContain(name, banned);
            Assert.False(name.StartsWith("Microsoft.AspNetCore", StringComparison.Ordinal), $"Cis.Application must not reference {name}");
        });
    }

    [Fact]
    public void Infrastructure_does_not_reference_the_api()
    {
        Assert.DoesNotContain("Cis.Api", References(Infrastructure));
        Assert.Contains("Cis.Application", References(Infrastructure));
    }

    [Fact]
    public void Controllers_use_the_application_layer_only()
    {
        var banned = new Regex(@"\b(Cis\.Infrastructure|Dapper|SqlConnection|SqlCommand|DbConnections|IDbSessions?|try|catch)\b");
        foreach (var file in Directory.GetFiles(Path.Combine(RepoPaths.CisRoot, "src", "Cis.Api", "Controllers"), "*.cs"))
            Assert.False(banned.IsMatch(File.ReadAllText(file)), $"{Path.GetFileName(file)} touches infrastructure or catches exceptions");
        foreach (var c in Controllers)
            foreach (var p in c.GetConstructors().SelectMany(ctor => ctor.GetParameters()))
                Assert.True(p.ParameterType.Assembly == Application, $"{c.Name} depends on {p.ParameterType.Name}; controllers take application services only");
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
    public void Application_services_depend_on_abstractions_only()
    {
        foreach (var service in Application.GetTypes().Where(t => t.IsClass && !t.IsAbstract && t.Name.EndsWith("Service", StringComparison.Ordinal)))
            foreach (var p in service.GetConstructors().SelectMany(c => c.GetParameters()))
            {
                var type = p.ParameterType;
                var ok = type.IsInterface || type == typeof(TimeProvider) ||
                    (type.IsGenericType && (type.GetGenericTypeDefinition() == typeof(ILogger<>) || type.GetGenericTypeDefinition() == typeof(IOptions<>)));
                Assert.True(ok, $"{service.Name} depends on concrete {type.Name}");
            }
    }

    [Fact]
    public void Repositories_live_in_infrastructure_behind_application_interfaces()
    {
        foreach (var repository in Infrastructure.GetTypes().Where(t => t.IsClass && !t.IsAbstract && t.Name.EndsWith("Repository", StringComparison.Ordinal)))
            Assert.Contains(repository.GetInterfaces(), i => i.Assembly == Application && i.Name == "I" + repository.Name);
        Assert.DoesNotContain(Application.GetTypes(), t => t.IsClass && t.Name.EndsWith("Repository", StringComparison.Ordinal));
    }

    [Fact]
    public void Every_application_port_is_registered()
    {
        using var scope = factory.Services.CreateScope();
        foreach (var contract in Application.GetTypes().Where(t => t.IsInterface && t.IsPublic && !t.IsGenericType && t.Name != "IDbSession"))
            Assert.True(scope.ServiceProvider.GetService(contract) is not null, $"{contract.Name} is not registered");
    }

    [Fact]
    public void Every_request_model_has_a_validator()
    {
        using var scope = factory.Services.CreateScope();
        var requests = Controllers.SelectMany(c => c.GetMethods()).SelectMany(m => m.GetParameters())
            .Select(p => p.ParameterType).Where(t => t.Name.EndsWith("Request", StringComparison.Ordinal)).Distinct().ToList();
        Assert.NotEmpty(requests);
        foreach (var request in requests)
        {
            Assert.Equal(Application, request.Assembly);
            Assert.True(scope.ServiceProvider.GetService(typeof(IValidator<>).MakeGenericType(request)) is not null, $"{request.Name} has no validator");
        }
    }

    [Fact]
    public void Every_screen_endpoint_outside_auth_audit_and_lookups_carries_an_access_check()
    {
        string[] shared = ["AuthController", "AuditController", "LookupsController"];
        foreach (var c in Controllers.Where(c => !shared.Contains(c.Name)))
        {
            var classLevel = c.GetCustomAttributes<RequireAccessAttribute>().Any();
            foreach (var action in c.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
                Assert.True(classLevel || action.GetCustomAttributes<RequireAccessAttribute>().Any(), $"{c.Name}.{action.Name} has no [RequireAccess]");
        }
    }
}
