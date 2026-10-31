using FluentValidation;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace Cis.Api.Common;

/// <summary>
/// Runs the registered <see cref="IValidator{T}"/> of every action argument before the action executes and answers
/// HTTP 400 with the first error. Controllers therefore never validate by hand.
/// </summary>
public sealed class ValidationFilter(IServiceProvider services) : IAsyncActionFilter
{
    /// <inheritdoc />
    public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
    {
        foreach (var argument in context.ActionArguments.Values)
        {
            if (argument is null) continue;
            if (services.GetService(typeof(IValidator<>).MakeGenericType(argument.GetType())) is not IValidator validator) continue;
            var result = await validator.ValidateAsync(new ValidationContext<object>(argument), context.HttpContext.RequestAborted);
            if (result.IsValid) continue;
            context.Result = ((ControllerBase)context.Controller).Invalid(result.Errors[0].ErrorMessage);
            return;
        }
        await next();
    }
}
