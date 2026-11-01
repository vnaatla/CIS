using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Membership;

/// <summary>The case and its current member (ListCases).</summary>
/// <param name="CasNumber">Case No.</param>
/// <param name="MbmNumber">Member No.</param>
/// <param name="CovId">Cov ID.</param>
/// <param name="Patient">Patient: the principal's name for cover 00, the covered person's otherwise.</param>
/// <param name="VisitDate">Visit Date.</param>
public record TransferCase(string CasNumber, string MbmNumber, string CovId, string Patient, DateTime? VisitDate);

/// <summary>Transfer of a case to another member.</summary>
/// <param name="CasNumber">Case No.</param>
/// <param name="MbmNumber">Transfer member number.</param>
/// <param name="CovId">Transfer cover id.</param>
public record CaseTransferRequest(string? CasNumber, string? MbmNumber, string? CovId);

/// <summary>Column lengths (CISCAS).</summary>
public sealed class CaseTransferRequestValidator : RequestValidator<CaseTransferRequest>
{
    /// <summary>Defines the rules.</summary>
    public CaseTransferRequestValidator()
    {
        RuleFor(r => r.CasNumber).MaximumLength(20);
        RuleFor(r => r.MbmNumber).MaximumLength(50);
        RuleFor(r => r.CovId).MaximumLength(2);
    }
}

/// <summary>Data access of Case Transfer (<c>cis.CaseTransfer_*</c>).</summary>
public interface ICaseTransferRepository
{
    /// <summary>The case's current member, or null.</summary>
    Task<TransferCase?> CaseAsync(IDbSession session, string casNumber);

    /// <summary>The member's (cover's) name, or null.</summary>
    Task<string?> MemberNameAsync(IDbSession session, string mbmNumber, string covId);

    /// <summary>Moves the case; rows updated.</summary>
    Task<int> SaveAsync(IDbSession session, string casNumber, string mbmNumber, string covId);
}

/// <summary>Invoice > Other > Case Transfer.</summary>
public interface ICaseTransferService
{
    /// <summary>CmdGo_Click.</summary>
    Task<TransferCase> FindAsync(string? casNumber);

    /// <summary>cmdSearch_Click: the transfer member's name.</summary>
    Task<string> MemberAsync(string? mbmNumber, string? covId);

    /// <summary>cmdSave_Click (Transfer).</summary>
    Task<string> SaveAsync(CaseTransferRequest request);
}

/// <inheritdoc />
public sealed class CaseTransferService(IDbSessions sessions, ICaseTransferRepository repository) : ICaseTransferService
{
    /// <summary>Desktop texts (spacing kept).</summary>
    public const string CaseFirst = "Please select Case Number ! ", NotFound = "Record not found!", MemberFirst = "Please enter Member Number! ",
        CoverFirst = "Please select Cover ID! ", Updated = "Record Updated!";

    /// <inheritdoc />
    public async Task<TransferCase> FindAsync(string? casNumber)
    {
        var cas = (casNumber ?? "").Trim();
        if (cas.Length == 0) throw new BusinessRuleException(CaseFirst);
        await using var s = await sessions.OpenAsync();
        return await repository.CaseAsync(s, cas) ?? throw new NotFoundException(NotFound);
    }

    /// <inheritdoc />
    public async Task<string> MemberAsync(string? mbmNumber, string? covId)
    {
        var (number, cov) = Check(mbmNumber, covId);
        await using var s = await sessions.OpenAsync();
        return await Name(s, number, cov);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(CaseTransferRequest r)
    {
        var cas = (r.CasNumber ?? "").Trim();
        if (cas.Length == 0) throw new BusinessRuleException(CaseFirst);
        var (number, cov) = Check(r.MbmNumber, r.CovId);
        await using var s = await sessions.BeginAsync();
        // Transfer is enabled only after Search found the member (cmdSearch_Click), and the case was loaded by Go.
        await Name(s, number, cov);
        if (await repository.SaveAsync(s, cas, number, cov) == 0) throw new NotFoundException(NotFound);
        await s.CommitAsync();
        return Updated;
    }

    private async Task<string> Name(IDbSession s, string number, string cov) =>
        await repository.MemberNameAsync(s, number, cov) is { Length: > 0 } name ? name : throw new NotFoundException(NotFound);

    private static (string Number, string Cov) Check(string? mbmNumber, string? covId)
    {
        if (string.IsNullOrWhiteSpace(mbmNumber)) throw new BusinessRuleException(MemberFirst);
        if (string.IsNullOrWhiteSpace(covId)) throw new BusinessRuleException(CoverFirst);
        return (mbmNumber.Trim(), covId.Trim());
    }
}
