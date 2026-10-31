using Cis.Application.Audit;
using Cis.Domain.Errors;
using Microsoft.AspNetCore.Diagnostics;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Data.SqlClient;

namespace Cis.Api.Common;

/// <summary>
/// The one place that turns exceptions into HTTP answers, so controllers and services carry no try/catch. Answers are Problem Details
/// with a <c>message</c> the user can act on. Failures the user cannot fix are logged with their stack trace, written to the legacy
/// <c>HISMaintenance.dbo.Tbl_ErrorLog</c> (SysName 'CIS', as the desktop's <c>Medix.ErrorLog</c> did) and answered with a generic message.
/// </summary>
public sealed class ApiExceptionHandler(ILogger<ApiExceptionHandler> log, IProblemDetailsService problemDetails, IErrorLog errorLog) : IExceptionHandler
{
    /// <summary>Status 499: the browser closed the request.</summary>
    public const int ClientClosedRequest = 499;

    /// <summary>The status, title and message an exception is answered with.</summary>
    public static (int Status, string Title, string Message) Map(Exception exception, bool requestAborted = false) => exception switch
    {
        NotFoundException => (StatusCodes.Status404NotFound, "Not found", exception.Message),
        ConflictException => (StatusCodes.Status409Conflict, "Conflict", exception.Message),
        UnauthorizedException => (StatusCodes.Status401Unauthorized, "Unauthorized", exception.Message),
        BusinessRuleException => (StatusCodes.Status400BadRequest, "Business rule", exception.Message),
        NotConfiguredException => (StatusCodes.Status503ServiceUnavailable, "Not configured", exception.Message),
        SqlException e when SqlErrorNumbers.IsTruncation(e.Number) =>
            (StatusCodes.Status400BadRequest, "Value too long", "One of the values is longer than the database allows. Nothing was saved."),
        SqlException e when SqlErrorNumbers.IsDuplicateKey(e.Number) =>
            (StatusCodes.Status409Conflict, "Duplicate", "This record already exists. Nothing was saved."),
        SqlException { Number: SqlErrorNumbers.Deadlock } =>
            (StatusCodes.Status409Conflict, "Deadlock", "Someone else is changing the same records. Wait a moment and try again."),
        SqlException { Number: SqlErrorNumbers.Timeout } =>
            (StatusCodes.Status504GatewayTimeout, "Database timeout", "The request took too long. Narrow the filters and try again."),
        SqlException { Number: SqlErrorNumbers.MissingProcedure } =>
            (StatusCodes.Status500InternalServerError, "Database not up to date",
                "The database is missing the stored procedures this version needs. Run Cis/sql/cis-main.sql on CISDB (see Cis/sql/README.md)."),
        OperationCanceledException when requestAborted => (ClientClosedRequest, "Request cancelled", "Stopped. Nothing was saved."),
        _ => (StatusCodes.Status500InternalServerError, "Server error", "Server error. Check the API log for details."),
    };

    /// <inheritdoc />
    public async ValueTask<bool> TryHandleAsync(HttpContext httpContext, Exception exception, CancellationToken cancellationToken)
    {
        var (status, title, message) = Map(exception, httpContext.RequestAborted.IsCancellationRequested);
        var path = httpContext.Features.Get<IExceptionHandlerPathFeature>()?.Path ?? httpContext.Request.Path.Value ?? "";
        if (status >= StatusCodes.Status500InternalServerError && status != StatusCodes.Status503ServiceUnavailable)
        {
            log.LogError(exception, "Unhandled error processing {Method} {Path}", httpContext.Request.Method, path);
            await errorLog.WriteAsync(path, httpContext.Request.Method, exception.Message, httpContext.User.Identity?.Name);
        }
        else
        {
            log.LogInformation("{Method} {Path} refused with {Status}: {Message}", httpContext.Request.Method, path, status, message);
        }

        httpContext.Response.StatusCode = status;
        return await problemDetails.TryWriteAsync(new ProblemDetailsContext
        {
            HttpContext = httpContext,
            Exception = exception,
            ProblemDetails = Problems.Create(httpContext, status, title, message),
        });
    }
}
