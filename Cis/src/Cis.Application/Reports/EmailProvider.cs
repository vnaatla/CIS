using System.Globalization;
using System.Text;
using Cis.Application.Abstractions;
using Cis.Application.Email;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Reports;

/// <summary>Email Notification (Providers) search (frmEmailNotificationProvider cmdSearch_Click).</summary>
public sealed record EmlPrvRequest(string? PvDateFrom, string? PvDateTo, string? PvFrom, string? PvTo, string? Status, int Entity);

/// <summary>Send: the ticked clinic PVs as listed, the banked-in date and extra BCC addresses.</summary>
public sealed record EmlPrvSendRequest(string? SendDate, string? ExtraBcc, List<EmlPrvTick> Ticked);

/// <summary>A ticked clinic PV as listed.</summary>
public sealed record EmlPrvTick(string PvNo, string? Email, string? ClinicCode, string? ClinicName, string? ChequeNo, string? Amount);

/// <summary>Lengths.</summary>
public sealed class EmlPrvRequestValidator : Common.RequestValidator<EmlPrvRequest>
{
    /// <summary>Defines the rules.</summary>
    public EmlPrvRequestValidator()
    {
        RuleFor(r => r.PvFrom).MaximumLength(30);
        RuleFor(r => r.PvTo).MaximumLength(30);
    }
}

/// <summary>Lengths.</summary>
public sealed class EmlPrvSendRequestValidator : Common.RequestValidator<EmlPrvSendRequest>
{
    /// <summary>Defines the rules.</summary>
    public EmlPrvSendRequestValidator()
    {
        RuleFor(r => r.ExtraBcc).MaximumLength(500);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 2000 && t.All(x => x.PvNo.Length <= 30 && (x.Email ?? "").Length <= 300));
    }
}

/// <summary>Data access.</summary>
public interface IEmlPrvRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, DateTime from, DateTime to, string? pvFrom, string? pvTo, bool resend, bool cs);
    Task<List<Dictionary<string, object?>>> InvoicesAsync(IDbSession session, string pvNo);
    Task MarkAsync(IDbSession session, string pvNo);
#pragma warning restore CS1591
}

/// <summary>Invoice > Email Notification (Providers) (frmEmailNotificationProvider).</summary>
public interface IEmlPrvService
{
    /// <summary>cmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(EmlPrvRequest request);

    /// <summary>CmdSend_Click (after the confirmation): the number of e-mails sent.</summary>
    Task<int> SendAsync(EmlPrvSendRequest request);
}

/// <inheritdoc />
public sealed class EmlPrvService(IDbSessions sessions, IEmlPrvRepository repository, IEmailSender mail) : IEmlPrvService
{
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;
    private const string FntNormal = "<font color=black face=\"Arial\" size=1>";

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(EmlPrvRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.PvDateFrom)) throw new BusinessRuleException(EmlPayService.NoFrom);
        if (string.IsNullOrWhiteSpace(r.PvDateTo)) throw new BusinessRuleException(EmlPayService.NoTo);
        if (string.IsNullOrWhiteSpace(r.Status)) throw new BusinessRuleException(EmlPayService.NoStatus);
        var from = ClinicInvoiceRules.Date(r.PvDateFrom) ?? throw new BusinessRuleException(EmlPayService.BadDate);
        var to = ClinicInvoiceRules.Date(r.PvDateTo) ?? throw new BusinessRuleException(EmlPayService.BadDate);
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, from, to, B(r.PvFrom), B(r.PvTo), r.Status.Trim().StartsWith('R'), r.Entity == 1);
        if (rows.Count == 0) throw new NotFoundException(EmlPayService.NoRecord);
        return rows;
    }

    /// <inheritdoc />
    public async Task<int> SendAsync(EmlPrvSendRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.SendDate)) throw new BusinessRuleException(EmlPayService.NoSendDate);
        if (ClinicInvoiceRules.Date(r.SendDate) is null) throw new BusinessRuleException(EmlPayService.BadDate);
        if (r.Ticked.Count == 0) throw new BusinessRuleException(EmlPayService.NoSelection);
        if (r.Ticked.Any(t => string.IsNullOrWhiteSpace(t.Email))) throw new BusinessRuleException(EmlPayService.NoEmail);
        var bcc = new List<string> { EmlPayService.Bcc };
        bcc.AddRange((r.ExtraBcc ?? "").Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
        var sent = 0;
        foreach (var t in r.Ticked)
        {
            List<Dictionary<string, object?>> invoices;
            await using (var s = await sessions.OpenAsync()) invoices = await repository.InvoicesAsync(s, t.PvNo.Trim());
            var (entity, body) = Body(t, invoices, r.SendDate.Trim());
            await mail.SendAsync(new EmailMessage([t.Email!.Trim()], $"{entity} Payment Advice - {t.ClinicName} ({t.ClinicCode})", body) { Bcc = bcc });
            await using (var s = await sessions.OpenAsync()) await repository.MarkAsync(s, t.PvNo.Trim());
            sent++;
        }
        return sent;
    }

    /// <summary>The desktop's HTML: the payment advice and the table of the PV's invoices with totals; the sender is MAJU SIHAT for entity SJ, else HEALTH CONNECT.</summary>
    public static (string Entity, string Html) Body(EmlPrvTick t, List<Dictionary<string, object?>> invoices, string sendDate)
    {
        const string nl = "\r\n";
        var rows = new StringBuilder();
        decimal bill = 0, paid = 0, copay = 0, ffs = 0;
        var n = 1;
        var entityCode = "";
        foreach (var d in invoices)
        {
            rows.Append("<tr><td>").Append(n++).Append("</td> <td>").Append(V(d, "InvoiceNo")).Append("</td> <td>").Append(V(d, "InvoiceDate")).Append("</td> <td>")
                .Append(V(d, "PVCLNDate")).Append("</td> <td>").Append(M(d, "CLMTotalActual")).Append("</td> <td>").Append(M(d, "CLMCoPayColl")).Append("</td> <td>")
                .Append(M(d, "CLMMCOProcFee")).Append("</td> <td>").Append(M(d, "CLMPayableByMedix2c")).Append("</td> </tr>");
            bill += R(d, "CLMTotalActual");
            paid += R(d, "CLMPayableByMedix2c");
            copay += R(d, "CLMCoPayColl");
            ffs += R(d, "CLMMCOProcFee");
            entityCode = V(d, "ENTITYCODE");
        }
        var entity = entityCode == "SJ" ? "MAJU SIHAT" : "HEALTH CONNECT";
        var table = "<table border=1>" + FntNormal +
            "<font size=2 face=\"Arial\"><b> <tr bgcolor=lightblue> <td>No</td> <td>Invoice No</td> <td>Invoice Date</td> <td>Cheque Date</td> <td>Bill Amount</td> <td>Co-Payment</td> <td>FFS</td> <td>Amount Paid</td></tr></b></font>" +
            rows + "<tr></tr> <tr><td>Total </td> <td>----</td> <td>----</td> <td>----</td> <td>" + bill.ToString("0.00", En) + "</td> <td>" + copay.ToString("0.00", En) +
            "</td> <td>" + ffs.ToString("0.00", En) + "</td> <td>" + paid.ToString("0.00", En) + "</td> </tr>" + "</font>" + "</table>";
        var amount = decimal.TryParse(t.Amount, NumberStyles.Any, En, out var a) ? a.ToString("0.00", En) : t.Amount;
        var body = "Dear Sir/Madam,</p></p> " + nl + nl +
            "PAYMENT ADVICE </p>  " + nl + nl +
            "We are pleased to advise that payment in respect of your invoice(s) as mentioned below has been banked into your Account. </p> Please allow reasonable time for the cheque to be cleared.</p>" + nl + nl +
            "PV Number               :       " + t.PvNo.Trim() + "</p>" +
            "Cheque Number         :       " + t.ChequeNo + "</p>" +
            "PV Amount               :       " + "RM " + amount + "</p>" +
            "Date Banked-in         :       " + sendDate + "</p>" +
            "DETAILS OF PAYMENT" + "</p>" +
            table + "</p>" +
            "</p>Please contact our following staff if you require any clarification:- </p> " +
            "           General Line    -  03-7884 1919 </p></p>" +
            "Thank you. </p>" +
            "Regards,</p>" +
            entity + " SDN BHD";
        return (entity, "<HTML><BODY>" + FntNormal + body + " </BODY></HTML>");
    }

    private static string V(IReadOnlyDictionary<string, object?> r, string k) => r.GetValueOrDefault(k) switch
    {
        null or DBNull => "",
        DateTime d => d.ToString("dd/MM/yyyy", En),
        var o => Convert.ToString(o, En)!.Trim(),
    };

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) => decimal.TryParse(V(r, k), NumberStyles.Any, En, out var d) ? d : 0;

    private static string M(IReadOnlyDictionary<string, object?> r, string k) => N(r, k).ToString("0.00", En);

    /// <summary>The desktop adds each amount as formatted to 2 decimals.</summary>
    private static decimal R(IReadOnlyDictionary<string, object?> r, string k) => Math.Round(N(r, k), 2, MidpointRounding.AwayFromZero);
}
