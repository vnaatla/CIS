using Cis.Api.Common;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Extensions.Options;
using MimeKit;

namespace Cis.Api.Email;

/// <summary>Outgoing mail settings (configuration section <c>Email</c>). Replaces the desktop's Outlook automation.</summary>
public class EmailOptions
{
    /// <summary>Configuration section name.</summary>
    public const string SectionName = "Email";

    /// <summary>SMTP server; empty = not configured (unless <see cref="PickupDirectory"/> is set).</summary>
    public string? Host { get; set; }

    /// <summary>SMTP port.</summary>
    public int Port { get; set; } = 587;

    /// <summary>Use STARTTLS / SSL when the server offers it.</summary>
    public bool UseTls { get; set; } = true;

    /// <summary>SMTP user (empty = anonymous).</summary>
    public string? User { get; set; }

    /// <summary>SMTP password (set it in an environment variable, never in appsettings.json).</summary>
    public string? Password { get; set; }

    /// <summary>Sender address.</summary>
    public string From { get; set; } = "notifications@medix.com.my";

    /// <summary>When set, messages are written as .eml files to this folder instead of being sent (development and tests).</summary>
    public string? PickupDirectory { get; set; }
}

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

/// <summary>Sends e-mail.</summary>
public interface IEmailSender
{
    /// <summary>Sends the message, or throws <see cref="NotConfiguredException"/> when e-mail is not set up.</summary>
    Task SendAsync(EmailMessage message, CancellationToken cancellationToken = default);
}

/// <summary><see cref="IEmailSender"/> with MailKit: SMTP, or .eml files in a pickup folder.</summary>
public sealed class SmtpEmailSender(IOptions<EmailOptions> options) : IEmailSender
{
    /// <inheritdoc />
    public async Task SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
    {
        var o = options.Value;
        var mime = Build(message, o.From);
        if (!string.IsNullOrWhiteSpace(o.PickupDirectory))
        {
            Directory.CreateDirectory(o.PickupDirectory);
            await mime.WriteToAsync(Path.Combine(o.PickupDirectory, $"{DateTime.UtcNow:yyyyMMddHHmmssfff}-{Guid.NewGuid():N}.eml"), cancellationToken);
            return;
        }
        if (string.IsNullOrWhiteSpace(o.Host))
            throw new NotConfiguredException("E-mail is not configured. Set Email:Host (or Email:PickupDirectory) on the API server.");

        using var smtp = new SmtpClient();
        await smtp.ConnectAsync(o.Host, o.Port, o.UseTls ? SecureSocketOptions.StartTlsWhenAvailable : SecureSocketOptions.None, cancellationToken);
        if (!string.IsNullOrWhiteSpace(o.User)) await smtp.AuthenticateAsync(o.User, o.Password ?? "", cancellationToken);
        await smtp.SendAsync(mime, cancellationToken);
        await smtp.DisconnectAsync(true, cancellationToken);
    }

    /// <summary>The MIME message for <paramref name="message"/>.</summary>
    public static MimeMessage Build(EmailMessage message, string from)
    {
        var mime = new MimeMessage();
        mime.From.Add(MailboxAddress.Parse(from));
        foreach (var to in message.To) mime.To.Add(MailboxAddress.Parse(to));
        foreach (var cc in message.Cc) mime.Cc.Add(MailboxAddress.Parse(cc));
        mime.Subject = message.Subject;
        var body = new BodyBuilder { HtmlBody = message.HtmlBody };
        foreach (var (fileName, content) in message.Attachments) body.Attachments.Add(fileName, content);
        mime.Body = body.ToMessageBody();
        return mime;
    }
}
