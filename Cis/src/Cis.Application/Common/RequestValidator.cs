using FluentValidation;

namespace Cis.Application.Common;

/// <summary>
/// Base class for request validators. The first failing rule is the one reported (the desktop forms showed one message at a
/// time, in a fixed order), so rules are declared in the order the user should fix them.
/// </summary>
public abstract class RequestValidator<T> : AbstractValidator<T>
{
    /// <summary>Stops at the first failing rule.</summary>
    protected RequestValidator()
    {
        ClassLevelCascadeMode = CascadeMode.Stop;
        RuleLevelCascadeMode = CascadeMode.Stop;
    }
}
