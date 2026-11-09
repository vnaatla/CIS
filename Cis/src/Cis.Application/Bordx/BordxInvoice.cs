using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Bordx;

/// <summary>Insert Invoice search (frmWSBordx2 SearchRecord).</summary>
public sealed record BordxInvSearchRequest(string? RegisterFrom, string? RegisterTo, string? Payor, string? GroupCompany);

/// <summary>Lengths.</summary>
public sealed class BordxInvSearchRequestValidator : Common.RequestValidator<BordxInvSearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxInvSearchRequestValidator()
    {
        RuleFor(r => r.Payor).MaximumLength(300);
        RuleFor(r => r.GroupCompany).MaximumLength(300);
    }
}

/// <summary>A ticked invoice row.</summary>
public sealed record BordxInvTick(string InvoiceNo, DateTime? InvoiceDate, DateTime? RegisterDate, string GrpCoCode);

/// <summary>Bordx: the ticked rows and the group company combo ("code - name").</summary>
public sealed record BordxInvSaveRequest(string? GroupCompany, List<BordxInvTick> Ticked);

/// <summary>Lengths.</summary>
public sealed class BordxInvSaveRequestValidator : Common.RequestValidator<BordxInvSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxInvSaveRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(300);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.InvoiceNo.Length <= 50 && x.GrpCoCode.Length <= 10));
    }
}

/// <summary>The listing and its total approved.</summary>
public record BordxInvResult(List<Dictionary<string, object?>> Rows, decimal TotalApproved);

/// <summary>Data access.</summary>
public interface IBordxInvRepository
{
#pragma warning disable CS1591
    Task<(List<(string Code, string Name, string? PayCode)> Companies, List<(string Code, string Name)> Payors)> LookupsAsync(IDbSession session);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, DateTime? from, DateTime? to, string? payCode, string? grpCode);
    Task<List<(string CasNumber, decimal? CLMFinalApprove)>> CasesAsync(IDbSession session, BordxInvTick tick);
    Task UpdateCaseAsync(IDbSession session, string casNumber, string bordxNo);
    Task SaveBordxAsync(IDbSession session, string bordxNo, string payCode, string company, string user, decimal amount, int claims);
#pragma warning restore CS1591
}

/// <summary>Claim Bordx > Worksheet > Insert Invoice (frmWSBordx2).</summary>
public interface IBordxInvService
{
    /// <summary>The combos.</summary>
    Task<object> LookupsAsync();

    /// <summary>CmdSearch_Click.</summary>
    Task<BordxInvResult> SearchAsync(BordxInvSearchRequest request);

    /// <summary>CmdBordx_Click (after the confirmation): the bordx no.</summary>
    Task<string> BordxAsync(BordxInvSaveRequest request, string user);
}

/// <inheritdoc />
public sealed class BordxInvService(IDbSessions sessions, IBordxInvRepository repository, IBordxWsRepository numbers) : IBordxInvService
{
    /// <summary>Desktop texts.</summary>
    public const string NoRecord = "No record found!", NoDate = "Invoice Date cannot be empty.", NoCompany = "Group Company cannot be empty!",
        EmptyList = "Please select a record before proceeding!", NoneTicked = "Please select at least one record to be updated! ", Saved = "Record Saved, Your Bordx No is : ";

    /// <inheritdoc />
    public async Task<object> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        var (companies, payors) = await repository.LookupsAsync(s);
        return new
        {
            companies = companies.Select(c => new { code = c.Code, name = c.Name, payCode = c.PayCode }),
            payors = payors.Select(p => new { code = p.Code, name = p.Name }),
        };
    }

    /// <inheritdoc />
    public async Task<BordxInvResult> SearchAsync(BordxInvSearchRequest r)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, ClinicInvoiceRules.Date(r.RegisterFrom), ClinicInvoiceRules.Date(r.RegisterTo), Code(r.Payor), Code(r.GroupCompany));
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        foreach (var x in rows)
            x["TickError"] = x.GetValueOrDefault("InvoiceDate") is null or DBNull ? NoDate
                : string.IsNullOrWhiteSpace(Convert.ToString(x.GetValueOrDefault("GRPCoName"), CultureInfo.InvariantCulture)) ? NoCompany : null;
        return new BordxInvResult(rows, rows.Sum(x => x.GetValueOrDefault("CLMTotalApproved") is { } v and not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0));
    }

    /// <inheritdoc />
    public async Task<string> BordxAsync(BordxInvSaveRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoneTicked);
        if (r.Ticked.Any(t => t.InvoiceDate is null)) throw new BusinessRuleException(NoDate);
        await using var tx = await sessions.BeginAsync();
        var (companies, _) = await repository.LookupsAsync(tx);
        // As the desktop, the bordx no. and BordxPAYCode come from the group company chosen in the combo.
        var payCode = companies.FirstOrDefault(c => c.Code.Equals(Code(r.GroupCompany) ?? "", StringComparison.OrdinalIgnoreCase)).PayCode?.Trim() ?? "";
        var bordxNo = await numbers.NextNoAsync(tx, payCode);
        decimal total = 0;
        var count = 0;
        foreach (var t in r.Ticked.Distinct())
            foreach (var (cas, approved) in await repository.CasesAsync(tx, t))
            {
                total += approved ?? 0;
                count++;
                await repository.UpdateCaseAsync(tx, cas, bordxNo);
            }
        await repository.SaveBordxAsync(tx, bordxNo, payCode, (r.GroupCompany ?? "").Trim(), user, total, count);
        await tx.CommitAsync();
        return bordxNo;
    }

    /// <summary>The code of "code - name".</summary>
    private static string? Code(string? v)
    {
        var t = (v ?? "").Trim();
        if (t.Length == 0) return null;
        var i = t.IndexOf(" - ", StringComparison.Ordinal);
        return (i < 0 ? t : t[..i]).Trim();
    }
}
