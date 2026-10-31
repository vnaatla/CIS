using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Claims;
using System.Text.RegularExpressions;
using Cis.Api.Auth;
using Cis.Api.Common;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Abstractions;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Routing;
using Microsoft.Data.SqlClient;

namespace Cis.Api.Tests;

public class AccessRightTests
{
    [Theory]
    [InlineData("101", 1, true)]
    [InlineData("101", 2, false)]
    [InlineData("101", 3, true)]
    [InlineData("101", 4, false)]
    [InlineData(null, 1, false)]
    [InlineData("1", 0, false)]
    public void Has_reads_position_n_of_usr_access(string? access, int position, bool expected) =>
        Assert.Equal(expected, AccessRight.Has(access, position));

    [Fact]
    public void Positions_come_from_the_cis_menu_handlers()
    {
        // Spot checks against MDIMain.frm: MBMAdjustment_Click CheckAccess(113), mnuINVAdjAfterBordx_Click 179, mnuCheckItem_Click 143.
        Assert.Equal(113, AccessRight.MbmAdjustment);
        Assert.Equal(179, AccessRight.InvoiceAdjustmentAfterBordx);
        Assert.Equal(143, AccessRight.MaintenanceCheckItem);
        Assert.Equal(120, AccessRight.ClaimTracking2);
    }

    [Fact]
    public void The_angular_access_table_matches_this_one()
    {
        var ts = File.ReadAllText(Path.Combine(RepoPaths.RepoRoot, "CisUI", "src", "app", "core", "auth", "access.ts"));
        var web = Regex.Matches(ts, @"^\s*(\w+):\s*(\d+),", RegexOptions.Multiline).ToDictionary(m => m.Groups[1].Value, m => int.Parse(m.Groups[2].Value));
        var api = typeof(AccessRight).GetFields(BindingFlags.Public | BindingFlags.Static).Where(f => f.IsLiteral)
            .ToDictionary(f => f.Name, f => (int)f.GetRawConstantValue()!);
        Assert.Equal(api.OrderBy(kv => kv.Key), web.OrderBy(kv => kv.Key));
    }

    private static AuthorizationFilterContext Context(string? access)
    {
        var identity = access is null ? new ClaimsIdentity() : new ClaimsIdentity([new Claim(TokenService.AccessClaim, access)], "test");
        var http = new DefaultHttpContext { User = new ClaimsPrincipal(identity) };
        return new AuthorizationFilterContext(new ActionContext(http, new RouteData(), new ActionDescriptor()), []);
    }

    [Fact]
    public void Require_access_answers_access_denied_without_the_position()
    {
        var context = Context(new string('1', 112) + "0");
        new RequireAccessAttribute(113).OnAuthorization(context);
        var result = Assert.IsType<ObjectResult>(context.Result);
        Assert.Equal(403, result.StatusCode);
        Assert.Equal("Access Denied!", Assert.IsType<ProblemDetails>(result.Value).Extensions["message"]);
    }

    [Fact]
    public void Require_access_passes_with_any_listed_position()
    {
        var context = Context(new string('0', 112) + "1");
        new RequireAccessAttribute(160, 113).OnAuthorization(context);
        Assert.Null(context.Result);
    }
}

public class ExceptionMappingTests
{
    [Fact]
    public void Business_exceptions_keep_their_message_and_status()
    {
        Assert.Equal((400, "Business rule", "x"), ApiExceptionHandler.Map(new BusinessRuleException("x")));
        Assert.Equal((404, "Not found", "x"), ApiExceptionHandler.Map(new NotFoundException("x")));
        Assert.Equal((409, "Conflict", "x"), ApiExceptionHandler.Map(new ConflictException("x")));
        Assert.Equal((401, "Unauthorized", "x"), ApiExceptionHandler.Map(new UnauthorizedException("x")));
        Assert.Equal(503, ApiExceptionHandler.Map(new NotConfiguredException("x")).Status);
    }

    [Fact]
    public void Unexpected_errors_hide_the_details()
    {
        var (status, _, message) = ApiExceptionHandler.Map(new InvalidOperationException("secret connection string"));
        Assert.Equal(500, status);
        Assert.DoesNotContain("secret", message);
    }

    [Fact]
    public void Cancelled_requests_are_499_only_when_the_browser_went_away()
    {
        Assert.Equal(499, ApiExceptionHandler.Map(new OperationCanceledException(), requestAborted: true).Status);
        Assert.Equal(500, ApiExceptionHandler.Map(new OperationCanceledException()).Status);
    }

    [Fact]
    public void Sql_errors_the_user_can_act_on_are_explained()
    {
        Assert.Equal(504, ApiExceptionHandler.Map(MakeSqlException(SqlErrorNumbers.Timeout)).Status);
        Assert.Equal(409, ApiExceptionHandler.Map(MakeSqlException(SqlErrorNumbers.Deadlock)).Status);
        Assert.Contains("cis-main.sql", ApiExceptionHandler.Map(MakeSqlException(SqlErrorNumbers.MissingProcedure)).Message);
    }

    /// <summary>SqlException has no public constructor; build one with the given error number (same trick as the GENESIS tests).</summary>
    private static SqlException MakeSqlException(int number)
    {
        var error = (SqlError)RuntimeHelpers.GetUninitializedObject(typeof(SqlError));
        typeof(SqlError).GetFields(BindingFlags.NonPublic | BindingFlags.Instance)
            .First(f => f.FieldType == typeof(int) && f.Name.Contains("number", StringComparison.OrdinalIgnoreCase)).SetValue(error, number);
        var errors = (SqlErrorCollection)Activator.CreateInstance(typeof(SqlErrorCollection), true)!;
        typeof(SqlErrorCollection).GetMethod("Add", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(errors, [error]);
        var exception = (SqlException)RuntimeHelpers.GetUninitializedObject(typeof(SqlException));
        typeof(SqlException).GetFields(BindingFlags.NonPublic | BindingFlags.Instance).First(f => f.FieldType == typeof(SqlErrorCollection)).SetValue(exception, errors);
        return exception;
    }
}
