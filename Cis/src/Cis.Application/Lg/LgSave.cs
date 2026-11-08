using System.Globalization;
using System.Net;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Email;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Lg;

/// <summary>Issue (cmdLG_Click + SaveLG) of the LG loaded and opened with Edit.</summary>
public sealed record LgIssueRequest
{
#pragma warning disable CS1591
    public string? LgNo { get; init; }
    public string? Member { get; init; }
    public string? Patient { get; init; }
    public string? Ic { get; init; }
    public string? ClinicText { get; init; }
    public string? GlType { get; init; }
    public string? Doctor { get; init; }
    public string? TreatmentDate { get; init; }
    public string? Remarks { get; init; }
    public string? InternalRemarks { get; init; }
    public string? Diagnosis { get; init; }
    public string? GlAmount { get; init; }
    public string? MaternityNormal { get; init; }
    public string? Caesarean { get; init; }
    public string? Complications { get; init; }
    public string? DelayReason { get; init; }
    public string? GroupCompany { get; init; }
    public string? Category { get; init; }
    public string? Subsidiary { get; init; }
    public string? Fax { get; init; }
    /// <summary>Edit was accepted for this LG (the desktop's EditFlag).</summary>
    public bool Edited { get; init; }
    /// <summary>Answers to the questions asked so far, in order (true = OK / Yes).</summary>
    public List<bool> Answers { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class LgIssueRequestValidator : Common.RequestValidator<LgIssueRequest>
{
    /// <summary>Defines the rules.</summary>
    public LgIssueRequestValidator()
    {
        RuleFor(r => r.Doctor).MaximumLength(200);
        RuleFor(r => r.ClinicText).MaximumLength(300);
        RuleFor(r => r.Diagnosis).MaximumLength(4000);
        RuleFor(r => r.Remarks).MaximumLength(8000);
        RuleFor(r => r.InternalRemarks).MaximumLength(4000);
        RuleFor(r => r.DelayReason).MaximumLength(500);
        RuleFor(r => r.Answers).Must(a => a.Count <= 10);
    }
}

/// <summary>Decline (cmdDecline_Click).</summary>
public sealed record LgDeclineRequest
{
#pragma warning disable CS1591
    public string? LgNo { get; init; }
    public string? Reason { get; init; }
    public string? Member { get; init; }
    public string? Patient { get; init; }
    public string? Ic { get; init; }
    public string? ClinicText { get; init; }
    public string? Doctor { get; init; }
    public string? Diagnosis { get; init; }
    public string? TreatmentDate { get; init; }
    public List<bool> Answers { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class LgDeclineRequestValidator : Common.RequestValidator<LgDeclineRequest>
{
    /// <summary>Defines the rules.</summary>
    public LgDeclineRequestValidator()
    {
        RuleFor(r => r.Reason).MaximumLength(500);
        RuleFor(r => r.Doctor).MaximumLength(200);
        RuleFor(r => r.Diagnosis).MaximumLength(4000);
        RuleFor(r => r.Answers).Must(a => a.Count <= 5);
    }
}

/// <summary>An action's outcome: a question to answer, a stop, or done with its messages and the letter.</summary>
/// <param name="Kind">Ok (done), Confirm (answer <paramref name="Message"/> and call again), Stop.</param>
/// <param name="Message">Question or stopping message.</param>
/// <param name="YesNo">The question is Yes / No (else OK / Cancel).</param>
/// <param name="Messages">Messages shown (in order) when done.</param>
/// <param name="Letter">The generated letter (PDF), if any.</param>
public record LgActionResult(LgGateKind Kind, string? Message, bool YesNo, List<string> Messages, FileOutput? Letter = null);

/// <summary>Results of cis.Lg_Checks.</summary>
public sealed record LgChecks
{
#pragma warning disable CS1591
    public bool HasClaim { get; init; }
    public bool Bordx { get; init; }
    public string InvoiceNo { get; init; } = "";
    public bool BtbgBordx { get; init; }
    public bool Paid { get; init; }
    public bool SpClinicOk { get; init; }
    public string? ClinicPanel { get; init; }
    public DateTime? Expiry { get; init; }
    public bool Pending { get; init; }
#pragma warning restore CS1591
}

/// <summary>cis.Lg_Issue parameters.</summary>
public record LgIssueSave(string LgNo, string User, DateTime TreatmentDate, string ClinicCode, string Remarks, string Diagnosis, decimal LgAmount, decimal GlAmount,
    string InternalRemarks, string Doctor, string GlType, decimal Maternity, bool Pending, bool Delivery, string DelayReason);

/// <summary>Writes of Generate LG.</summary>
public interface ILgSaveRepository
{
#pragma warning disable CS1591
    Task<LgChecks> ChecksAsync(IDbSession session, string lgNo, string member, string cover, string clinicCode, string clinicText);
    Task IssueAsync(IDbSession session, LgIssueSave save);
    Task DeclineAsync(IDbSession session, string lgNo, string doctor, string clinicCode, string reason, string user, string diagnosis, bool pharmacy);
    Task<(string To, string Bcc, string Cc, string Subsidiary)?> EmailRecipientsAsync(IDbSession session, string company, string member);
#pragma warning restore CS1591
}

public sealed partial class LgService
{
    /// <summary>Desktop texts.</summary>
    public const string SelectLg = "Please Select the LG !!", Bordxed = "This case cannot amend already bordx Or Invoice no already inserted !!",
        Paid = "THE CASE CANNOT AMEND ALREADY PAYMENT DONE !!", NotPanelSp = "THE SELECTED CLINIC NOT IN PANEL !!",
        NotHcPanel = "The selected clinic is not in HC panel. GL not allowed..!!", Expired = "THE ABOVE MEMBER EXPIRY..!!", PressEdit = "Kindly Press Edit..Then Continue",
        Insufficient = "Insufficient Balance, Please Check", Updated = "Record Updated....", Generated = "LG Generated Successfully",
        DeclineFirst = "Please Select Decline Reason", SelectLgDecline = "Please Select a LG!", Declined = "Decline Letter Generated Successfully",
        Copied = "Specified File Copied Successfully", SelectRecord = "Please Select Record!",
        BadRxType = "Invalid File Type. Please upload a File with extension (bmp, gif, png, jpg, jpeg, pdf)";

    private static readonly string[] Deliveries = ["Inpatient Delivery Private", "Inpatient Delivery Private-Complications", "Inpatient Delivery (GOV)"];
    private static readonly string[] NoAmountTypes = ["Pre-Employee Checkup", "Inpatient Delivery Private", "Inpatient Delivery Private-Complications", "Inpatient Delivery (GOV)"];

    private sealed class Ask(string message, bool yesNo) : Exception(message) { public bool YesNo { get; } = yesNo; }

    private sealed class Halt(string? message) : Exception(message ?? "") { public string? Text { get; } = message; }

    private sealed class Answers(List<bool> answers)
    {
        private int _i;

        /// <summary>The answer to the next question (throws to ask it).</summary>
        public bool Next(string question, bool yesNo) => _i < answers.Count ? answers[_i++] : throw new Ask(question, yesNo);
    }

    private static async Task<LgActionResult> Run(Func<List<string>, Task<FileOutput?>> body)
    {
        var said = new List<string>();
        try
        {
            var letter = await body(said);
            return new LgActionResult(LgGateKind.Ok, null, false, said, letter);
        }
        catch (Ask a) { return new LgActionResult(LgGateKind.Confirm, a.Message, a.YesNo, said); }
        catch (Halt h) { return new LgActionResult(LgGateKind.Stop, h.Text, false, said); }
    }

    /// <summary>CmdAdd_Click (after a cover was clicked): the limit checks; on Ok the page opens the entry fields.</summary>
    public async Task<LgGateResult> AddCheckAsync(string? member, string? cover, int confirmed)
    {
        var info = await MemberAsync(member, cover);
        if (info.Status != "A") return new LgGateResult(LgGateKind.Stop, NotActive, []);
        return LgGates.Add(info.Limits, info.ValidLg, true, info.Member.PolicyExpiry, clock.GetLocalNow().DateTime, info.Member.GroupCompany, confirmed);
    }

    /// <summary>CmdEdit_Click: the case / clinic / expiry checks, then the limit checks (with the LG's own labels, as lstLG_Click set them).</summary>
    public async Task<LgGateResult> EditCheckAsync(string? lgNo, string? clinicText, string? category, int confirmed)
    {
        if (string.IsNullOrWhiteSpace(lgNo)) return new LgGateResult(LgGateKind.Stop, SelectLg, []);
        var lg = Case(lgNo);
        var (member, cover, _, info) = await LgMember(lg);
        await using var s = await sessions.OpenAsync();
        var c = await saves.ChecksAsync(s, lg, member, cover, ClinicCode(clinicText), (clinicText ?? "").Trim());
        if (c.HasClaim && (c.Bordx || c.InvoiceNo.Length > 0 || c.BtbgBordx)) return new LgGateResult(LgGateKind.Stop, Bordxed, []);
        if (c.Paid) return new LgGateResult(LgGateKind.Stop, Paid, []);
        var pemc = (category ?? "").Trim() == "PEMC";
        if (!pemc && !c.SpClinicOk) return new LgGateResult(LgGateKind.Stop, NotPanelSp, []);
        if (!pemc && c.ClinicPanel is null or "N" or "D") return new LgGateResult(LgGateKind.Stop, NotHcPanel, []);
        if (c.Expiry is null) return new LgGateResult(LgGateKind.Stop, Expired, []);
        var raw = (await repository.MemberAsync(s, member, cover, "")).Limits ?? [];
        return LgGates.Edit(LgRules.RawLimits(raw, info.Limits, info.Member.Aso), info.GpSpShared, info.Member.GroupCompany, confirmed);
    }

    /// <summary>cmdLG_Click + SaveLG: issues the LG loaded with Edit, writes the GL letter and sends the company's notification.</summary>
    public Task<LgActionResult> IssueAsync(LgIssueRequest r, string user) => Run(async said =>
    {
        var ask = new Answers(r.Answers);
        if (string.IsNullOrWhiteSpace(r.LgNo)) throw new Halt(SelectLg);
        var lg = Case(r.LgNo);
        var (member, cover, item, info) = await LgMember(lg);
        var clinicCode = ClinicCode(r.ClinicText);
        await using var s = await sessions.OpenAsync();
        var c = await saves.ChecksAsync(s, lg, member, cover, clinicCode, (r.ClinicText ?? "").Trim());
        if (c.Pending && !ask.Next(PendingNotice, false)) throw new Halt(null);
        var raw = LgRules.RawLimits((await repository.MemberAsync(s, member, cover, "")).Limits ?? [], info.Limits, info.Member.Aso);
        var previous = LgRules.Num(T(item, "GLAmt"));
        var valid = BalanceValid(r.GlAmount, raw, info.GpSpShared, info.AfterLimitExInd, (r.GlType ?? "").Trim(), previous);
        if (c.HasClaim && (c.Bordx || c.InvoiceNo.Length > 0 || c.BtbgBordx)) throw new Halt(Bordxed);
        if (c.Paid) throw new Halt(Paid);
        var pemc = (r.Category ?? "").Trim() == "PEMC";
        if (!pemc && c.ClinicPanel is null or "N" or "D") throw new Halt(NotHcPanel);
        if (!pemc && !c.SpClinicOk) throw new Halt(NotPanelSp);
        if (c.Expiry is null) throw new Halt(Expired);
        var now = clock.GetLocalNow().DateTime;
        var date = ClinicInvoiceRules.Date(r.TreatmentDate);
        if (T(item, "LGTIME") is { Length: > 0 } caseTime && item.GetValueOrDefault("LGDate") is DateTime caseDate && date is { } d0 && caseDate.Date == d0.Date
            && TimeSpan.TryParse(caseTime, CultureInfo.InvariantCulture, out var t0) && (now.TimeOfDay - t0).TotalMinutes > 30)
        {
            if (string.IsNullOrWhiteSpace(r.DelayReason)) throw new Halt("Please Enter the reason for Delay.!");
            if (ask.Next("Are you want to Enter or Update the reason for Delay.?", true)) throw new Halt(null);
        }
        if (!valid) throw new Halt(Insufficient);
        return await SaveLg(r, user, s, ask, said, lg, member, cover, clinicCode, info, c, date, now);
    });

    private async Task<FileOutput?> SaveLg(LgIssueRequest r, string user, Abstractions.IDbSession s, Answers ask, List<string> said, string lg, string member, string cover,
        string clinicCode, LgMemberInfo info, LgChecks c, DateTime? date, DateTime now)
    {
        var glType = (r.GlType ?? "").Trim();
        if (string.IsNullOrWhiteSpace(r.ClinicText)) throw new Halt("Please select a Clinic");
        if (glType.Length == 0) throw new Halt("Please select Type of GL");
        if (string.IsNullOrWhiteSpace(r.Member)) throw new Halt("Please select a Patient");
        if (string.IsNullOrWhiteSpace(r.Patient)) throw new Halt("Please select a Patient Name");
        if (string.IsNullOrWhiteSpace(r.Ic)) throw new Halt("Please select Patient's IC Number");
        if (string.IsNullOrWhiteSpace(r.Doctor)) throw new Halt("Please Enter the Treating Doctor's Name");
        if (string.IsNullOrWhiteSpace(r.GlAmount) && !NoAmountTypes.Contains(glType)) throw new Halt("Please Enter the GL Amount");
        if (info.Status != "A") throw new Halt(NotActive);
        if (string.IsNullOrWhiteSpace(r.TreatmentDate)) throw new Halt("Please enter the Treatment Date.");
        if (date is null) throw new Halt("Invalid Date Format, please re-enter the Treatment Date.");
        if (string.IsNullOrWhiteSpace(r.Diagnosis)) throw new Halt("Please Enter the Diagnosis");
        if (c.Expiry < date) throw new Halt("Treatment Date cannot be Greater Than Member Expiry date..!");
        var template = await repository.TemplateAsync(s, clinicCode, (r.GroupCompany ?? "").Trim(), info.Limits.PlanCode, glType, "", (r.Category ?? "").Trim(), lg);
        var templateName = cover != "00" && template.DependantInd == "Y" ? template.DependantTemplate : template.TemplateName;
        if (templateName.Length == 0) throw new Halt($"Please Check NO GLTemplate is Assigned for Plan Code : {info.Limits.PlanCode}. Please Check LG Maintainace Table");
        if (!ask.Next("Are you sure?", true)) throw new Halt(null);
        if (c.Paid) throw new Halt(null);
        if (!r.Edited) throw new Halt(PressEdit);

        string normal = (r.MaternityNormal ?? "").Trim() is { Length: > 0 } n ? n : "0", caes = (r.Caesarean ?? "").Trim() is { Length: > 0 } cs ? cs : "0",
            comp = (r.Complications ?? "").Trim() is { Length: > 0 } cp ? cp : "0";
        // The desktop compares the two text boxes as text.
        var maternityText = glType == "Inpatient Delivery Private" ? (string.CompareOrdinal(normal, caes) >= 0 ? caes : normal)
            : glType == "Inpatient Delivery Private-Complications" ? comp : "0";
        var maternity = LgRules.Num(maternityText);
        var gl = LgRules.Num(r.GlAmount);
        var privateDelivery = glType is "Inpatient Delivery Private" or "Inpatient Delivery Private-Complications";
        await saves.IssueAsync(s, new LgIssueSave(lg, user, date!.Value, clinicCode, r.Remarks ?? "", (r.Diagnosis ?? "").Trim(), privateDelivery ? maternity : gl, gl,
            r.InternalRemarks ?? "", (r.Doctor ?? "").Trim(), glType, privateDelivery ? maternity : 0, c.Pending, Deliveries.Contains(glType), (r.DelayReason ?? "").Trim()));
        said.Add(Updated);

        var pdf = letters.Write(GlLetter(r, lg, user, now, template, privateDelivery ? maternity : gl));
        var file = new FileOutput(pdf, "application/pdf", $"L{lg}.pdf");
        await SaveLetter(lg, file);
        said.Add(Generated);

        var lgs = await repository.ListAsync(s, member, cover);
        if (lgs.Count > 0 && T(lgs[^1], "EmailInd") != "1"
            && await saves.EmailRecipientsAsync(s, (r.GroupCompany ?? "").Trim(), member) is { To.Length: > 0 } to)
        {
            var body = GlEmail(r, info, lg, to.Subsidiary);
            try
            {
                await mail.SendAsync(new EmailMessage([to.To], $"Notification of Outpatient LG Issued To : {r.Patient} , Membership Number :{r.Member}",
                    "<pre>" + WebUtility.HtmlEncode(body) + "</pre>")
                {
                    Cc = Split(to.Cc), Bcc = Split(to.Bcc), Attachments = [(file.FileName, pdf)],
                });
            }
            catch (NotConfiguredException e) { said.Add(e.Message); }
        }
        return file;
    }

    /// <summary>cmdLG_Click's balance test (tmpIsValid). Desktop quirks kept: a blank amount is a VB type mismatch (nothing happens), and with a
    /// previous GL amount the test compares IsNumeric(amount) (-1 / 0) with the balance.</summary>
    internal static bool BalanceValid(string? amountText, LgLimits l, string gpSp, string after, string glType, decimal previous)
    {
        var amount = (amountText ?? "").Trim();
        if (!LgRules.IsNum(amount)) throw new Halt(null);
        var gl = LgRules.Num(amount);
        if (!(gl >= 0 && after == "N")) return true;
        var valid = false;
        decimal bal = 0;
        var edit = previous != 0;
        var cmp = edit ? -1m : gl;
        // Edit path: (tmpBal + balance) + tmpGLAmt_Chk_update for each numeric balance added.
        if (gpSp == "Y")
        {
            if (LgRules.IsNum(l.SpecBalance)) bal += LgRules.Num(l.SpecBalance) + (edit ? previous : 0); else valid = true;
            if (LgRules.IsNum(l.Balance)) bal += LgRules.Num(l.Balance) + (edit ? previous : 0); else valid = true;
            if (l.Balance.Trim() == "FULL") valid = true;
            else if (edit) valid = cmp <= bal;
            else if (glType == "Normal GL") valid = gl <= bal;
        }
        else if (gpSp == "N")
        {
            if (LgRules.IsNum(l.SpecBalance)) bal += LgRules.Num(l.SpecBalance) + (edit ? previous : 0); else valid = true;
            valid = l.SpecBalance.Trim() == "FULL" || cmp <= bal;
        }
        return valid;
    }

    /// <summary>cmdDecline_Click: declines the LG and writes the decline letter.</summary>
    public Task<LgActionResult> DeclineAsync(LgDeclineRequest r, string user) => Run(async said =>
    {
        var ask = new Answers(r.Answers);
        if (string.IsNullOrWhiteSpace(r.Reason)) throw new Halt(DeclineFirst);
        if (string.IsNullOrWhiteSpace(r.LgNo)) throw new Halt(SelectLgDecline);
        var lg = Case(r.LgNo);
        if (!ask.Next($"Are you sure you want to decline [{lg}]?", true)) throw new Halt(null);
        var (_, _, _, info) = await LgMember(lg);
        await using var s = await sessions.OpenAsync();
        var d = await repository.LoadAsync(s, lg);
        await saves.DeclineAsync(s, lg, (r.Doctor ?? "").Trim(), ClinicCode(r.ClinicText), r.Reason.Trim(), user, (r.Diagnosis ?? "").Trim(), d.Pharmacy);
        var now = clock.GetLocalNow().DateTime;
        var pdf = letters.Write(new LetterDocument("DECLINE OF GUARANTEE LETTER", [now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)],
        [
            ("Case Number", lg), ("Date", now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)), ("Patient Name", Proper(r.Patient)),
            ("Employee", info.Member.Name), ("Membership No", (r.Member ?? "").Trim()), ("IC Number", (r.Ic ?? "").Trim()), ("Hospital", (r.ClinicText ?? "").Trim()),
            ("Admission Date", (r.TreatmentDate ?? "").Trim()), ("Treatment Date", (r.TreatmentDate ?? "").Trim()), ("Decline Reason", r.Reason.Trim()), ("Issued By", user),
        ], []));
        var file = new FileOutput(pdf, "application/pdf", $"D{lg}.pdf");
        await SaveLetter(lg, file);
        said.Add(Declined);
        return file;
    });

    /// <summary>CmdPreview_Click: the GL as it stands (no save).</summary>
    public async Task<FileOutput> PreviewAsync(LgIssueRequest r, string user)
    {
        var lg = Case(r.LgNo);
        var (_, _, _, info) = await LgMember(lg);
        await using var s = await sessions.OpenAsync();
        var template = await repository.TemplateAsync(s, ClinicCode(r.ClinicText), (r.GroupCompany ?? "").Trim(), info.Limits.PlanCode, (r.GlType ?? "").Trim(), "",
            (r.Category ?? "").Trim(), lg);
        if (template.ClinicName.Length == 0) throw new BusinessRuleException("Please check GL format!");
        return new FileOutput(letters.Write(GlLetter(r, lg, user, clock.GetLocalNow().DateTime, template, LgRules.Num(r.GlAmount))), "application/pdf", $"Preview{lg}.pdf");
    }

    /// <summary>cmdCopyFile_Click / CopyFile: copies a file into the LG's folder (pharmacy cases: RX_&lt;LG&gt;-&lt;n&gt; in OtherScannedDoc).</summary>
    public async Task<string> CopyAsync(string? lgNo, string? fileName, Stream content)
    {
        if (string.IsNullOrWhiteSpace(lgNo)) throw new BusinessRuleException(SelectRecord);
        var lg = Case(lgNo);
        await using var s = await sessions.OpenAsync();
        var pharmacy = (await repository.LoadAsync(s, lg)).Pharmacy;
        var dir = Folder(lg, pharmacy) ?? throw new NotConfiguredException("The outpatient documents folder is not configured (Files:Folders:Outpatient).");
        var name = Path.GetFileName((fileName ?? "").Trim());
        if (name.Length == 0 || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) throw new BusinessRuleException("Please choose a file.");
        var ext = Path.GetExtension(name).TrimStart('.');
        Directory.CreateDirectory(dir);
        string target;
        if (pharmacy)
        {
            if (!new[] { "bmp", "gif", "png", "jpg", "jpeg", "pdf" }.Contains(ext)) throw new BusinessRuleException(BadRxType);
            var n = Directory.EnumerateFiles(dir).Count(f => Path.GetFileName(f).StartsWith("RX_", StringComparison.Ordinal));
            target = Path.Combine(dir, $"RX_{lg}-{n}.{ext}");
            if (File.Exists(target)) throw new BusinessRuleException($"The file '{Path.GetFileName(target)}' already exists.");
        }
        else target = Path.Combine(dir, name);
        await using (var f = File.Create(target)) await content.CopyToAsync(f);
        return Copied;
    }

    /// <summary>A document of the LG's folder.</summary>
    public async Task<(byte[] Content, string FileName)> DocumentAsync(string? lgNo, string? fileName)
    {
        var lg = Case(lgNo);
        await using var s = await sessions.OpenAsync();
        var pharmacy = (await repository.LoadAsync(s, lg)).Pharmacy;
        var name = Documents(lg, pharmacy).FirstOrDefault(f => string.Equals(f, fileName?.Trim(), StringComparison.OrdinalIgnoreCase))
            ?? throw new NotFoundException("File not found.");
        return (await File.ReadAllBytesAsync(Path.Combine(Folder(lg, pharmacy)!, name)), name);
    }

    private async Task SaveLetter(string lg, FileOutput file)
    {
        if (Folder(lg, false) is not { } dir) return;
        Directory.CreateDirectory(dir);
        await File.WriteAllBytesAsync(Path.Combine(dir, file.FileName), file.Content);
    }

    /// <summary>The LG's member / cover, its LGlst2020 row and the member click data.</summary>
    private async Task<(string Member, string Cover, Dictionary<string, object?> Item, LgMemberInfo Info)> LgMember(string lg)
    {
        string member, cover;
        await using (var s = await sessions.OpenAsync())
        {
            var d = await repository.LoadAsync(s, lg);
            var row = d.Rows.FirstOrDefault(x => T(x, "CoverID") == T(x, "MBMCOVID")) ?? d.Rows.FirstOrDefault(x => T(x, "CoverID") == T(x, "MBMCCoverID"))
                ?? throw new NotFoundException(NotFound);
            member = T(row, "MBMNumber");
            cover = T(row, "CoverID");
        }
        var info = await MemberAsync(member, cover);
        var item = info.Lgs.FirstOrDefault(x => T(x, "LGNo") == lg) ?? throw new NotFoundException(NotFound);
        return (member, cover, item, info);
    }

    private static LetterDocument GlLetter(LgIssueRequest r, string lg, string user, DateTime now, LgTemplate t, decimal amount)
    {
        var company = t.GlFormat == "M" ? "MediExpress (Malaysia) Sdn Bhd" : "Health Connect (Malaysia) Sdn Bhd";
        var fields = new List<(string, string)>
        {
            ("LG No", $"{lg}-{user}"), ("Date", now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)), ("Patient", Proper(r.Patient)), ("IC No", (r.Ic ?? "").Trim()),
            ("Membership No", (r.Member ?? "").Trim()), ("Group Company", (r.GroupCompany ?? "").Trim()), ("Subsidiary", (r.Subsidiary ?? "").Trim()),
            ("Hospital", (r.ClinicText ?? "").Trim()), ("Address", string.Join("\n", new[] { t.ClinicAddress1, t.ClinicAddress2, t.ClinicState }.Where(x => x.Length > 0))),
            ("Fax", (r.Fax ?? "").Trim()), ("Treatment Date", (r.TreatmentDate ?? "").Trim()), ("Attending Doctor", (r.Doctor ?? "").Trim()),
            ("Diagnosis", (r.Diagnosis ?? "").Trim()), ("Type of GL", (r.GlType ?? "").Trim()), ("GL Amount (RM)", LgRules.Fmt(amount.ToString(CultureInfo.InvariantCulture))),
        };
        if (LgRules.IsNum(r.MaternityNormal)) fields.Add(("Normal Delivery (RM)", r.MaternityNormal!.Trim()));
        if (LgRules.IsNum(r.Caesarean)) fields.Add(("Caesarean (RM)", r.Caesarean!.Trim()));
        if (LgRules.IsNum(r.Complications)) fields.Add(("Complications (RM)", r.Complications!.Trim()));
        return new LetterDocument("GUARANTEE LETTER", [company, $"Template: {(t.TemplateName.Length > 0 ? t.TemplateName : "-")}"], fields,
            (r.Remarks ?? "").Replace("\r\n", "\n").Split('\n').Where(x => x.Trim().Length > 0).ToList());
    }

    private static string GlEmail(LgIssueRequest r, LgMemberInfo info, string lg, string subsidiary)
    {
        const string nl = "\r\n\r\n";
        return "Dear Sir/Madam," + nl + "LG ISSUED NOTIFICATION : " + nl +
            "Please be advised that the above-mentioned member is an Outpatient treatment (Check-up) at below hospital." + nl +
            "Below is the case details for your kind attention." + nl +
            "Patient Name             :       " + r.Patient + nl + "Employee Name            :       " + info.Member.Name + nl +
            "IC Number                :       " + r.Ic + nl + "Group Company            :       " + r.GroupCompany + nl +
            "Subsidiary               :       " + subsidiary + nl + "LG Number                :       " + lg + nl +
            "Treatment Date           :       " + r.TreatmentDate + nl + "Hospital                 :       " + r.ClinicText + nl + " " + nl +
            "Please contact us at 603-7884 1919 if you require any further clarification." + nl + " " + nl + "Thank you." + nl + "Regards, " + nl +
            "Health Connect (Malaysia) Sdn Bhd" + nl;
    }

    private static List<string> Split(string v) => v.Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

    private static string Proper(string? v) => CultureInfo.InvariantCulture.TextInfo.ToTitleCase((v ?? "").Trim().ToLowerInvariant());

    /// <summary>The clinic code: the text inside "[...]" (as SaveLG takes it).</summary>
    internal static string ClinicCode(string? text)
    {
        var t = (text ?? "").Trim();
        var i = t.IndexOf('[');
        var j = t.LastIndexOf(']');
        return i >= 0 && j > i ? t[(i + 1)..j].Trim() : "";
    }
}
