using Cis.Api.Common;
using Cis.Api.Data;
using FluentValidation;

namespace Cis.Api.Audit;

/// <summary>A screen was opened (desktop <c>Form_Load</c> → <c>ActionLogUpdate(12, "frmName")</c>).</summary>
/// <param name="Screen">Legacy form name, e.g. <c>frmMBMCLNEnq</c>.</param>
public record ScreenOpenedRequest(string Screen);

/// <summary>Validates <see cref="ScreenOpenedRequest"/>.</summary>
public sealed class ScreenOpenedRequestValidator : RequestValidator<ScreenOpenedRequest>
{
    /// <summary>Longest screen name accepted.</summary>
    public const int MaxLength = 50;

    /// <summary>Defines the rules.</summary>
    public ScreenOpenedRequestValidator()
    {
        RuleFor(r => r.Screen).NotEmpty().WithMessage("Screen is required.")
            .MaximumLength(MaxLength).WithMessage($"Screen must be at most {MaxLength} characters.")
            .Matches("^[A-Za-z0-9_]+$").WithMessage("Screen must be a form name (letters, digits, underscore).");
    }
}

/// <summary>Data access of the action log.</summary>
public interface IActionLogRepository
{
    /// <summary>Inserts a <c>HISDB.dbo.tblActionLog</c> row (ProjType 'CIS').</summary>
    Task InsertAsync(IDbSession session, string userCode, int actionId, string notes);
}

/// <inheritdoc />
public sealed class ActionLogRepository : IActionLogRepository
{
    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, string userCode, int actionId, string notes) =>
        session.ExecAsync("cis.ActionLog_Insert", new { userId = userCode, actionId, notes });
}

/// <summary>Records user actions in the legacy action log.</summary>
public interface IActionLogService
{
    /// <summary>Records that <paramref name="userCode"/> opened <paramref name="screen"/> (action 12, as every desktop form did).</summary>
    Task ScreenOpenedAsync(string userCode, string screen);
}

/// <inheritdoc />
public sealed class ActionLogService(IDbSessions sessions, IActionLogRepository repository) : IActionLogService
{
    /// <summary>The only action id the CIS desktop wrote (130 forms call <c>ActionLogUpdate(12, ...)</c>).</summary>
    public const int ScreenOpened = 12;

    /// <inheritdoc />
    public async Task ScreenOpenedAsync(string userCode, string screen)
    {
        await using var session = await sessions.BeginAsync();
        await repository.InsertAsync(session, userCode, ScreenOpened, screen);
        await session.CommitAsync();
    }
}
