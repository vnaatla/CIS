using Cis.Domain.Errors;

namespace Cis.Application.Email;

/// <summary>A message to send.</summary>
/// <param name="To">Recipients.</param>
/// <param name="Subject">Subject line.</param>
/// <param name="HtmlBody">HTML body.</param>
public record EmailMessage(IReadOnlyList<string> To, string Subject, string HtmlBody)
{
    /// <summary>Copy recipients.</summary>
    public IReadOnlyList<string> Cc { get; init; } = [];

    /// <summary>Attachments (file name, bytes).</summary>
    public IReadOnlyList<(string FileName, byte[] Content)> Attachments { get; init; } = [];
}

/// <summary>Sends e-mail (replaces the desktop's Outlook automation).</summary>
public interface IEmailSender
{
    /// <summary>Sends the message, or throws <see cref="NotConfiguredException"/> when e-mail is not set up.</summary>
    Task SendAsync(EmailMessage message, CancellationToken cancellationToken = default);
}
