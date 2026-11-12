using System.Globalization;
using System.IO.Compression;
using System.Net;
using System.Text;
using Cis.Application.Abstractions;
using Cis.Application.Email;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Reports;

/// <summary>Email Notification (Rejection) search (frmEmailRejectedClaim cmdSearch_Click): register date range, case no. range, status, group company.</summary>
public sealed record EmlRejRequest(string? RegisterFrom, string? RegisterTo, string? CaseFrom, string? CaseTo, string? Status, string? GroupCompany);

/// <summary>A ticked rejected claim as listed.</summary>
public sealed record EmlRejTick(string CaseNo, string? Email, string? MemberName, string? PatientName, string? Ic, string? Policy, string? InvoiceNo, string? MemberNo,
    string? Submitted, string? VisitDate, string? Amount, string? Reason, string? Action, string? GroupCompany);

/// <summary>Send: the ticked claims.</summary>
public sealed record EmlRejSendRequest(List<EmlRejTick> Ticked);

/// <summary>The e-mails sent, and the letters written for BSN members without a personal e-mail (a zip, as the desktop's Word documents).</summary>
public sealed record EmlRejSendResult(int Sent, int Letters, string? LettersFileName, byte[]? LettersZip);

/// <summary>Lengths.</summary>
public sealed class EmlRejRequestValidator : Common.RequestValidator<EmlRejRequest>
{
    /// <summary>Defines the rules.</summary>
    public EmlRejRequestValidator()
    {
        RuleFor(r => r.CaseFrom).MaximumLength(20);
        RuleFor(r => r.CaseTo).MaximumLength(20);
        RuleFor(r => r.GroupCompany).MaximumLength(200);
    }
}

/// <summary>Lengths.</summary>
public sealed class EmlRejSendRequestValidator : Common.RequestValidator<EmlRejSendRequest>
{
    /// <summary>Defines the rules.</summary>
    public EmlRejSendRequestValidator() =>
        RuleFor(r => r.Ticked).Must(t => t.Count <= 2000 && t.All(x => x.CaseNo.Length <= 20 && (x.Email ?? "").Length <= 300));
}

/// <summary>Data access.</summary>
public interface IEmlRejRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, DateTime from, DateTime to, string? caseFrom, string? caseTo, bool resend, string? company);
    Task MarkAsync(IDbSession session, string casNumber);
#pragma warning restore CS1591
}

/// <summary>Invoice > Email Notification (Rejection) (frmEmailRejectedClaim).</summary>
public interface IEmlRejService
{
    /// <summary>cmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(EmlRejRequest request);

    /// <summary>CmdSend_Click (after the confirmation).</summary>
    Task<EmlRejSendResult> SendAsync(EmlRejSendRequest request);
}

/// <inheritdoc />
public sealed class EmlRejService(IDbSessions sessions, IEmlRejRepository repository, IEmailSender mail, TimeProvider clock) : IEmlRejService
{
    /// <summary>Desktop texts.</summary>
    public const string NoFrom = "Please Enter Registration Date From!", NoTo = "Please Enter Registration Date To!";

    /// <summary>The desktop's fixed blind copy.</summary>
    public const string Bcc = "foo@medix.com.my";

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(EmlRejRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.RegisterFrom)) throw new BusinessRuleException(NoFrom);
        if (string.IsNullOrWhiteSpace(r.RegisterTo)) throw new BusinessRuleException(NoTo);
        if (string.IsNullOrWhiteSpace(r.Status)) throw new BusinessRuleException(EmlPayService.NoStatus);
        var from = ClinicInvoiceRules.Date(r.RegisterFrom) ?? throw new BusinessRuleException(EmlPayService.BadDate);
        var to = ClinicInvoiceRules.Date(r.RegisterTo) ?? throw new BusinessRuleException(EmlPayService.BadDate);
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, from, to, B(r.CaseFrom), B(r.CaseTo), r.Status.Trim().StartsWith('R'), B(r.GroupCompany));
        if (rows.Count == 0) throw new NotFoundException(EmlPayService.NoRecord);
        return rows;
    }

    /// <summary>BSN members with no e-mail or a "bsn..." e-mail get a letter instead of an e-mail, as the desktop.</summary>
    public static bool IsLetter(EmlRejTick t)
    {
        var email = (t.Email ?? "").Trim();
        return (t.GroupCompany ?? "").Trim() == "BANK SIMPANAN NASIONAL" && (email.Length == 0 || email[..Math.Min(3, email.Length)].Trim() == "bsn");
    }

    /// <inheritdoc />
    public async Task<EmlRejSendResult> SendAsync(EmlRejSendRequest r)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(EmlPayService.NoSelection);
        var letters = new List<(string Name, string Text)>();
        var sent = 0;
        foreach (var t in r.Ticked)
        {
            if (IsLetter(t))
            {
                var member = (t.MemberNo ?? "").Contains('*') ? t.MemberNo![..t.MemberNo!.IndexOf('*')] : t.MemberNo ?? "";
                letters.Add(($"{member} _ {(t.MemberName ?? "").Trim()}.txt", "\r\n\r\n" + Text(t, false)));
            }
            else
            {
                if (string.IsNullOrWhiteSpace(t.Email)) throw new BusinessRuleException(EmlPayService.NoEmail);
                var html = "<HTML><BODY>" + WebUtility.HtmlEncode(Text(t, true)).Replace("\r\n", "<br>") + "</BODY></HTML>";
                await mail.SendAsync(new EmailMessage([t.Email.Trim()], $"CLAIM - {t.PatientName}(Ref No : {t.CaseNo.Trim()})", html) { Bcc = [Bcc] });
                sent++;
            }
            await using var s = await sessions.OpenAsync();
            await repository.MarkAsync(s, t.CaseNo.Trim());
        }
        if (letters.Count == 0) return new EmlRejSendResult(sent, 0, null, null);
        using var zip = new MemoryStream();
        using (var a = new ZipArchive(zip, ZipArchiveMode.Create, true))
            foreach (var (name, text) in letters)
            {
                using var w = new StreamWriter(a.CreateEntry(string.Concat(name.Split(Path.GetInvalidFileNameChars()))).Open(), Encoding.UTF8);
                w.Write(text);
            }
        return new EmlRejSendResult(sent, letters.Count, $"RejectedMails {clock.GetLocalNow():yyyyMMdd}.zip", zip.ToArray());
    }

    /// <summary>The desktop's text (the e-mail ends with the regards, the BSN letter does not).</summary>
    public static string Text(EmlRejTick t, bool email)
    {
        const string nl = "\r\n";
        var amount = decimal.TryParse(t.Amount, NumberStyles.Any, CultureInfo.InvariantCulture, out var a) ? a.ToString("0.00", CultureInfo.InvariantCulture) : t.Amount;
        var text = "Dear Sir/Madam" + nl + (email ? nl : "") +
            "CLAIM DETAILS" + nl +
            "Member Name            :" + (t.MemberName ?? "").Trim() + nl +
            "Policy Number          :" + (t.Policy ?? "").Trim() + nl +
            "IC Number              :" + (t.Ic ?? "").Trim() + nl +
            "Claim No               :" + t.CaseNo.Trim() + nl +
            "InvoiceNo              :" + (t.InvoiceNo ?? "").Trim() + nl +
            "Membership No          :" + t.MemberNo + nl +
            "Date Submitted         :" + t.Submitted + nl +
            "Visit Date             :" + t.VisitDate + nl +
            "Incurred Amount        :" + "RM " + amount + nl + nl +
            "Thank you for submitting the above claim to us. Please see the following reasons(s) and action(s) required(if any) :-" + nl +
            "Reason(s) - " + t.Reason + nl + nl +
            "Action(s) - " + t.Action + nl + nl +
            "All original documents will be retunred to your HR department within Fourteen(14) working days" + nl + nl +
            "Kindly reply within 30 days or claim will be rejected and will not be entertained." + nl + nl +
            "Please contact our following staff if you require any clarification:- " + nl +
            "           General Line    -  03-7884 1919" + nl + nl +
            "Thank you.";
        return email ? text + nl + "Regards, " + nl + "MediExpress (Malaysia) Sdn Bhd" : text;
    }
}
