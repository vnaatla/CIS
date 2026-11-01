using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Membership;

/// <summary>The member shown by MBM Topup Adj (FindMBM).</summary>
/// <param name="MbmNumber">Membership number.</param>
/// <param name="CovId">Cover id ("00" = principal).</param>
/// <param name="Name">Name (the covered person's when CovId is not 00).</param>
/// <param name="IcNo">IC No.</param>
/// <param name="Plan">Plan.</param>
/// <param name="InsType">Insured type.</param>
/// <param name="EffDate">Policy effective date.</param>
/// <param name="ExpDate">Policy expiry date.</param>
/// <param name="Status">"A - Active", "C - Cancelled", "Delete - Delete" or the code.</param>
/// <param name="GroupCompany">Group company (MBMOthers).</param>
/// <param name="Department">Department (MBMOthers).</param>
/// <param name="BalanceLimit">Balance limit (GP or SP balance of spUtilisationOP_test1; the dependant's MBMCCLNBalLmt otherwise).</param>
/// <param name="LimitCode">Supplementary limit status (AnnualLimit; N when none).</param>
/// <param name="LimitLabel">lblLimitType ("" unless insured type F / H or code N).</param>
/// <param name="TopupEnabled">False for code N (unlimited).</param>
public record TopupMember(string MbmNumber, string CovId, string Name, string IcNo, string Plan, string InsType, DateTime? EffDate, DateTime? ExpDate,
    string Status, string GroupCompany, string Department, decimal? BalanceLimit, string LimitCode, string LimitLabel, bool TopupEnabled);

/// <summary>Save of MBM Topup Adj.</summary>
/// <param name="Member">Member box as typed ("NUMBER-COVID").</param>
/// <param name="Gp">GP limit option (true) or SP limit (false).</param>
/// <param name="IcNo">IC No box.</param>
/// <param name="PrevIcNo">IC shown when the member was loaded.</param>
/// <param name="BalanceLimit">Balance Limit box.</param>
/// <param name="TopupAmt">Topup Amount box.</param>
/// <param name="LimitCode">Limit code loaded with the member.</param>
public record TopupSaveRequest(string? Member, bool Gp, string? IcNo, string? PrevIcNo, string? BalanceLimit, string? TopupAmt, string? LimitCode);

/// <summary>Box formats.</summary>
public sealed class TopupSaveRequestValidator : RequestValidator<TopupSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public TopupSaveRequestValidator()
    {
        RuleFor(r => r.IcNo).MaximumLength(30);
        RuleFor(r => r.Member).MaximumLength(30);
    }
}

/// <summary>Data access of MBM Topup Adj.</summary>
public interface IMbmTopupRepository
{
    /// <summary>HISDB SearchCISMemberProc rows (principal, then one per covered person).</summary>
    Task<List<Dictionary<string, object?>>> MemberAsync(IDbSession session, string mbmNumber);

    /// <summary>MBMOthers group company and department.</summary>
    Task<(string GroupCompany, string Department)> OthersAsync(IDbSession session, string mbmNumber);

    /// <summary>spUtilisationOP_test1 rows.</summary>
    Task<List<Dictionary<string, object?>>> UtilisationAsync(IDbSession session, string mbmNumber);

    /// <summary>Supplementary limit status and description (null when the plan has no annual limit).</summary>
    Task<(string Code, string? Desc)?> LimitTypeAsync(IDbSession session, string plan, string insType, DateTime effDate);

    /// <summary>Next / previous clinic member number.</summary>
    Task<string?> StepAsync(IDbSession session, string mbmNumber, bool next);

    /// <summary>History, IC change and new limit.</summary>
    Task SaveAsync(IDbSession session, TopupSave save, string userCode, DateTime today);
}

/// <summary>Values written by CmdSave_Click.</summary>
/// <param name="MbmNumber">MBMNumber.</param>
/// <param name="CovId">Cover id.</param>
/// <param name="Gp">GP option.</param>
/// <param name="IcNo">New IC.</param>
/// <param name="PrevIcNo">Previous IC.</param>
/// <param name="PrevLimit">Balance before.</param>
/// <param name="TopupAmt">Top-up.</param>
/// <param name="NewLimit">Balance after.</param>
/// <param name="LimitCode">Supplementary limit status.</param>
public record TopupSave(string MbmNumber, string CovId, bool Gp, string IcNo, string PrevIcNo, decimal PrevLimit, decimal TopupAmt, decimal NewLimit, string LimitCode);

/// <summary>Membership > MBM Topup Adj.</summary>
public interface IMbmTopupService
{
    /// <summary>CmdGo_Click → FindMBM (GP or SP balance).</summary>
    Task<TopupMember> FindAsync(string? member, bool gp);

    /// <summary>CmdNext_Click / CmdPreview_Click: the member number to show ("NUMBER- 00").</summary>
    Task<string> StepAsync(string? member, bool next);

    /// <summary>CmdSave_Click.</summary>
    Task<string> SaveAsync(TopupSaveRequest request, string userCode);
}

/// <inheritdoc />
public sealed class MbmTopupService(IDbSessions sessions, IMbmTopupRepository repository, TimeProvider clock) : IMbmTopupService
{
    /// <summary>Desktop texts (spacing kept).</summary>
    public const string EnterMember = "Please Enter Membership No", NotFound = "No record found!", Nothing = "Nothing to Adjust ! ",
        Updated = "Updated Successfully", Last = " Last Record Reached! ", First = "First Record Reached! \nNot a valid Membership Number Format! ";

    /// <inheritdoc />
    public async Task<TopupMember> FindAsync(string? member, bool gp)
    {
        var (number, cov) = Split(member);
        if (number.Length == 0) throw new BusinessRuleException(EnterMember);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.MemberAsync(s, number);
        if (rows.Count == 0) throw new NotFoundException(NotFound);
        var (group, dept) = await repository.OthersAsync(s, number);
        var p = rows[0];
        string name = T(p, "MBMName"), ic = T(p, "MBMIcBcPp"), plan = T(p, "PLNType"), status = T(p, "MBMCLNStatus");
        decimal? balance = null;
        if (cov != "00" && rows.FirstOrDefault(r => T(r, "MBMCCoverID") == cov) is { } c)
        {
            if (T(c, "MBMCName").Length > 0) name = T(c, "MBMCName");
            if (T(c, "MBMCICBCPPNo").Length > 0) ic = T(c, "MBMCICBCPPNo");
            if (T(c, "MBMCCLNPLNCode").Length > 0) plan = T(c, "MBMCCLNPLNCode");
            if (Num(c.GetValueOrDefault("MBMCCLNBalLmt")) is { } bal) balance = bal;
            if (T(c, "MBMCCLNStatus").Length > 0) status = T(c, "MBMCCLNStatus");
        }
        var ins = T(p, "INSType");
        var eff = p.GetValueOrDefault("MBMPolEffDate") as DateTime?;
        string code = "N", label = "";
        if (name.Length > 0)
        {
            // GetSptLmtInd
            var type = eff is null ? null : await repository.LimitTypeAsync(s, plan, ins, eff.Value);
            if (type is { } t) code = t.Code;
            if (ins is "F" or "H") label = type?.Desc is { Length: > 0 } d ? $"{code} - {d}" : code;
            if (code == "N") label = "N - Unlimited Limit";
            balance = Balance(await repository.UtilisationAsync(s, number), cov, gp) ?? balance;
        }
        return new TopupMember(number, cov, name, ic, plan, ins, eff, p.GetValueOrDefault("MBMPolExpDate") as DateTime?, StatusLabel(status), group, dept,
            balance, code, label, code != "N");
    }

    /// <inheritdoc />
    public async Task<string> StepAsync(string? member, bool next)
    {
        var (number, _) = Split(member);
        await using var s = await sessions.OpenAsync();
        return await repository.StepAsync(s, number, next) is { } found
            ? $"{found}- 00"
            : throw new NotFoundException(next ? Last : First);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(TopupSaveRequest r, string userCode)
    {
        var (number, cov) = Split(r.Member);
        if (number.Length == 0) throw new BusinessRuleException(EnterMember);
        var prev = Num(r.BalanceLimit) ?? 0;
        var topup = Num(r.TopupAmt) ?? 0;
        var ic = (r.IcNo ?? "").Trim().ToUpperInvariant();
        var prevIc = (r.PrevIcNo ?? "").Trim();
        if (topup == 0 && prevIc == ic) throw new BusinessRuleException(Nothing);
        await using var s = await sessions.BeginAsync();
        await repository.SaveAsync(s, new TopupSave(number, cov, r.Gp, ic, prevIc, prev, topup, prev + topup, (r.LimitCode ?? "N").Trim()), userCode,
            clock.GetLocalNow().Date);
        await s.CommitAsync();
        return Updated;
    }

    /// <summary>CmdGo_Click: "NUMBER-COVID" → number and cover id (00 without a dash).</summary>
    public static (string Number, string CovId) Split(string? text)
    {
        var t = (text ?? "").Trim().ToUpperInvariant();
        var at = t.IndexOf('-');
        if (at < 0) return (t, "00");
        return (at == 0 ? t : t[..at].Trim(), t[(at + 1)..].Trim());
    }

    /// <summary>optGPlmt / optSPlmt: GPBalance or SPBalance of the cover's utilisation row (0 when not numeric).</summary>
    public static decimal? Balance(IEnumerable<Dictionary<string, object?>> utilisation, string cov, bool gp) =>
        utilisation.FirstOrDefault(u => T(u, "CovId") == cov) is { } row ? Num(row.GetValueOrDefault(gp ? "GPBalance" : "SPBalance")) ?? 0 : null;

    /// <summary>FindMBM's status texts.</summary>
    public static string StatusLabel(string code) => code switch { "A" => "A - Active", "C" => "C - Cancelled", "D" => "Delete - Delete", _ => code };

    private static string T(Dictionary<string, object?> row, string key) =>
        row.TryGetValue(key, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal? Num(object? v) => v switch
    {
        null => null,
        decimal d => d,
        double d => (decimal)d,
        int i => i,
        string s when decimal.TryParse(s.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) => d,
        _ => null,
    };
}
