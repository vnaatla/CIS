using System.Globalization;
using System.Text;
using Cis.Application.Abstractions;
using Cis.Application.Email;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Reports;

/// <summary>Email Notification (Payment) search (frmEmailNotification cmdSearch_Click).</summary>
public sealed record EmlPayRequest(string? PvDateFrom, string? PvDateTo, string? PvFrom, string? PvTo, string? Status, string? GroupCompany, int Entity);

/// <summary>Send: the ticked PVs (with the e-mail, name, member no., bank account and submitted date as listed), the banked-in date and extra BCC addresses.</summary>
public sealed record EmlPaySendRequest(string? SendDate, string? ExtraBcc, List<EmlPayTick> Ticked);

/// <summary>A ticked PV as listed.</summary>
public sealed record EmlPayTick(string PvNo, string? Email, string? Name, string? MemberNo, string? BankAccount, string? Submitted);

/// <summary>Lengths.</summary>
public sealed class EmlPayRequestValidator : Common.RequestValidator<EmlPayRequest>
{
    /// <summary>Defines the rules.</summary>
    public EmlPayRequestValidator()
    {
        RuleFor(r => r.PvFrom).MaximumLength(30);
        RuleFor(r => r.PvTo).MaximumLength(30);
        RuleFor(r => r.GroupCompany).MaximumLength(200);
    }
}

/// <summary>Lengths.</summary>
public sealed class EmlPaySendRequestValidator : Common.RequestValidator<EmlPaySendRequest>
{
    /// <summary>Defines the rules.</summary>
    public EmlPaySendRequestValidator()
    {
        RuleFor(r => r.ExtraBcc).MaximumLength(500);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 2000 && t.All(x => x.PvNo.Length <= 30 && (x.Email ?? "").Length <= 300));
    }
}

/// <summary>Data access.</summary>
public interface IEmlPayRepository
{
#pragma warning disable CS1591
    Task<List<string>> CompaniesAsync(IDbSession session);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, DateTime from, DateTime to, string? pvFrom, string? pvTo, bool resend, string? company, bool cs);
    Task<List<Dictionary<string, object?>>> DetailsAsync(IDbSession session, string pvNo);
    Task MarkAsync(IDbSession session, string pvNo);
#pragma warning restore CS1591
}

/// <summary>Invoice > Email Notification (Payment) (frmEmailNotification).</summary>
public interface IEmlPayService
{
    /// <summary>The group company combo.</summary>
    Task<List<string>> CompaniesAsync();

    /// <summary>cmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(EmlPayRequest request);

    /// <summary>CmdSend_Click (after the confirmation): the number of e-mails sent.</summary>
    Task<int> SendAsync(EmlPaySendRequest request);
}

/// <inheritdoc />
public sealed class EmlPayService(IDbSessions sessions, IEmlPayRepository repository, IEmailSender mail) : IEmlPayService
{
    /// <summary>Desktop texts.</summary>
    public const string NoFrom = "Please Enter PV Date From!", NoTo = "Please Enter PV Date To!", NoStatus = "Please Enter Status!", NoRecord = "No record found !",
        BadDate = "Invalid date format!", NoSendDate = "Please enter send date", NoSelection = "Please select the record", NoEmail = "Please check member's email!",
        Sent = "Emails have been sent out successfully!";

    /// <summary>The desktop's fixed blind copy.</summary>
    public const string Bcc = "elaineyeow@medix.com.my";

    private static readonly CultureInfo En = CultureInfo.InvariantCulture;
    private const string FntNormal = "<font color=black face=\"Arial\" size=1>";

    /// <inheritdoc />
    public async Task<List<string>> CompaniesAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.CompaniesAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(EmlPayRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.PvDateFrom)) throw new BusinessRuleException(NoFrom);
        if (string.IsNullOrWhiteSpace(r.PvDateTo)) throw new BusinessRuleException(NoTo);
        if (string.IsNullOrWhiteSpace(r.Status)) throw new BusinessRuleException(NoStatus);
        var from = ClinicInvoiceRules.Date(r.PvDateFrom) ?? throw new BusinessRuleException(BadDate);
        var to = ClinicInvoiceRules.Date(r.PvDateTo) ?? throw new BusinessRuleException(BadDate);
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, from, to, B(r.PvFrom), B(r.PvTo), r.Status.Trim().StartsWith('R'), B(r.GroupCompany), r.Entity == 1);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return rows;
    }

    /// <inheritdoc />
    public async Task<int> SendAsync(EmlPaySendRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.SendDate)) throw new BusinessRuleException(NoSendDate);
        if (ClinicInvoiceRules.Date(r.SendDate) is null) throw new BusinessRuleException(BadDate);
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoSelection);
        if (r.Ticked.Any(t => string.IsNullOrWhiteSpace(t.Email))) throw new BusinessRuleException(NoEmail);
        var bcc = new List<string> { Bcc };
        bcc.AddRange((r.ExtraBcc ?? "").Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
        var sent = 0;
        foreach (var t in r.Ticked)
        {
            List<Dictionary<string, object?>> details;
            await using (var s = await sessions.OpenAsync()) details = await repository.DetailsAsync(s, t.PvNo.Trim());
            await mail.SendAsync(new EmailMessage([t.Email!.Trim()], $"{t.Name} ({t.MemberNo}) - {t.PvNo.Trim()}", Body(t, details, r.SendDate.Trim())) { Bcc = bcc });
            // The desktop marked the PV before sending; here it is marked once the e-mail has gone.
            await using (var s = await sessions.OpenAsync()) await repository.MarkAsync(s, t.PvNo.Trim());
            sent++;
        }
        return sent;
    }

    /// <summary>The desktop's HTML: the payment advice and the table of the PV's claims with totals.</summary>
    public static string Body(EmlPayTick t, List<Dictionary<string, object?>> details, string sendDate)
    {
        const string nl = "\r\n";
        var table = new StringBuilder("<table border=1>").Append(FntNormal)
            .Append("<font size=2 face=\"Arial\"><b> <tr bgcolor=lightblue><td>No</td> <td>Claim Ref.</td> <td>DOV</td> <td>Incurred Amt</td> <td>Paid Amt</td> <td>Claimant</td> <td>Company Name</td> </tr></b></font>");
        decimal actual = 0, paid = 0;
        var n = 1;
        foreach (var d in details)
        {
            actual += N(d, "CLMTotalActual");
            paid += N(d, "CLMBordxAmt");
            table.Append("<tr><td>").Append(n++).Append("</td> <td>").Append(V(d, "CasNumber")).Append("</td> <td>").Append(V(d, "DateVisit")).Append("</td> <td>")
                .Append(N(d, "CLMTotalActual").ToString("0.00", En)).Append("</td> <td>").Append(N(d, "CLMBordxAmt").ToString("0.00", En)).Append("</td><td>")
                .Append(V(d, "PatientName")).Append("</td><td>").Append(V(d, "MBMGroupCompany")).Append("</td></tr>");
        }
        table.Append("<tr></tr> <tr><td>Total </td> <td>----</td> <td>----</td><td>").Append(actual.ToString("0.00", En)).Append("</td> <td>")
            .Append(paid.ToString("0.00", En)).Append("</td><td>---------</td><td>---------</td></tr>").Append("</font>").Append("</table>");
        var body = "Dear Sir/Madam,</p></p>" + nl + nl +
            "EB PAYMENT ADVICE </p></p>" + nl + nl +
            "We are pleased to advise that payment in respect of your claim as mentioned below has been banked into your Account. </p></p> Please allow reasonable time for the cheque to be cleared. </p></p>" + nl + nl +
            "PV Number               :       " + t.PvNo.Trim() + "</p></p>" +
            "Date Submitted         :       " + t.Submitted + "</p></p>" +
            "Incurred Amount      :       " + "RM " + actual.ToString("0.00", En) + "</p></p>" +
            "Paid Amount             :       " + "RM " + paid.ToString("0.00", En) + "</p></p>" +
            "Date Banked-in         :       " + sendDate + "</p></p>" +
            "Bank Account No.     :       " + t.BankAccount + " <br> </p></p>" + nl +
            " <br> DETAILS OF PAYMENT" + " <br>  </p></p> <br> " +
            table + nl + "</p></p>" +
            " <br><br> For your information, the difference(if any) between the paid amount and the amount incurred is due to MEMBER'S SHARE / CO-PAYMENT." + "</p></p>" + nl +
            "</p> <br> <br> Please contact our following staff if you require any clarification:- </p></p> " + nl +
            "           HelpLine    -  03-7884 1919 </p></p>" + nl + nl +
            "</p> <br> Thank you.</p></p>" + nl +
            "</p> <br> Regards, </p></p> " + nl +
            "MediExpress (Malaysia) Sdn Bhd";
        return "<HTML><BODY>" + FntNormal + body + " </BODY></HTML>";
    }

    private static string V(IReadOnlyDictionary<string, object?> r, string k) => r.GetValueOrDefault(k) switch
    {
        null or DBNull => "",
        DateTime d => d.ToString("dd/MM/yyyy", En),
        var o => Convert.ToString(o, En)!.Trim(),
    };

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) => decimal.TryParse(V(r, k), NumberStyles.Any, En, out var d) ? d : 0;
}
