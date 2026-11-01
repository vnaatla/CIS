using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Membership;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Membership;

/// <summary><see cref="IMbmAdjustmentRepository"/> on <c>cis.MbmAdj_*</c>, HISDB SearchCISMemberProc / SearchCISMBMHistory and the shared member procedures.</summary>
public sealed class MbmAdjustmentRepository : IMbmAdjustmentRepository
{
    /// <inheritdoc />
    public async Task<AdjSetup> SetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.MbmAdj_Setup");
        var cards = (await grid.ReadAsync<string>()).ToList();
        var rels = (await grid.ReadAsync<LookupOption>()).ToList();
        return new AdjSetup(cards, rels);
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> MemberAsync(IDbSession session, string mbmNumber) =>
        session.RowsAsync("cis.Hdb_SearchCISMemberProc", new { MBMNumber = mbmNumber });

    /// <inheritdoc />
    public async Task<(string GroupCompany, string Department)> OthersAsync(IDbSession session, string mbmNumber)
    {
        var row = (await session.RowsAsync("cis.MbmTopup_Others", new { mbmNumber })).FirstOrDefault();
        return (row?["GroupCompany"] as string ?? "", row?["Department"] as string ?? "");
    }

    /// <inheritdoc />
    public async Task<string> LimitStatusAsync(IDbSession session, string plan, string insType) =>
        (await session.ScalarAsync<string>("cis.MbmAdj_LimitStatus", new { plnCode = plan, insCode = insType }))?.Trim() ?? "N";

    /// <inheritdoc />
    public async Task<string> CoverLimitStatusAsync(IDbSession session, string plan, string insType, DateTime effDate) =>
        (await session.ScalarAsync<string>("cis.MbmAdj_CoverLimitStatus", new { plnCode = plan, insCode = insType, effDate }))?.Trim() ?? "A";

    /// <inheritdoc />
    public async Task<FeeLookup> PremiumAsync(IDbSession session, string mbmType, string insType, string payCode, string ageCode, string plan, DateTime effDate) =>
        (await session.QueryAsync<FeeLookup>("cis.MbmAdj_Premium", new { mbmType, insCode = insType, payCode, ageCode, plnCode = plan, hltCode = "N", effDate })).Last();

    /// <inheritdoc />
    public async Task<FeeLookup> McoAsync(IDbSession session, string mbmType, string insType, string adultChild, DateTime effDate, string plan, string payCode) =>
        (await session.QueryAsync<FeeLookup>("cis.MbmAdj_Mco", new { mbmType, insCode = insType, hltCode = "N", adultChild, effDate, plnCode = plan, payCode })).Single();

    /// <inheritdoc />
    public async Task<LimitLookup> LimitsAsync(IDbSession session, string mbmType, string insType, string plan, DateTime effDate, string payCode, string limitStatus) =>
        (await session.QueryAsync<LimitLookup>("cis.MbmAdj_Limits", new { mbmType, insCode = insType, plnCode = plan, effDate, payCode, limitStatus })).Single();

    /// <inheritdoc />
    public async Task<string> AgeCodeAsync(IDbSession session, int age, bool days)
    {
        var code = (await session.RowsAsync("cis.MbmReg_AgeBand", new { age, days })).FirstOrDefault()?["AgeCode"] as string;
        return int.TryParse(code, out var n) ? n.ToString("00", System.Globalization.CultureInfo.InvariantCulture) : "00";
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, (decimal PremReduction, decimal PremEffective, decimal McoRefund, decimal McoEffective)>> RefundsAsync(IDbSession session, string mbmNumber) =>
        (await session.QueryAsync<(string CoverId, decimal? PremReduction, decimal? PremEffective, decimal? McoRefund, decimal? McoEffective)>("cis.MbmAdj_Refunds", new { mbmNumber }))
            .GroupBy(r => r.CoverId)
            .ToDictionary(g => g.Key, g => (g.Last().PremReduction ?? 0, g.Last().PremEffective ?? 0, g.Last().McoRefund ?? 0, g.Last().McoEffective ?? 0));

    /// <inheritdoc />
    public async Task<string?> StepAsync(IDbSession session, string mbmNumber, bool next) =>
        (await session.RowsAsync("cis.MbmTopup_Step", new { mbmNumber, next })).FirstOrDefault()?["MBMNumber"] as string;

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> AmendHistoryAsync(IDbSession session, string mbmNumber) =>
        session.RowsAsync("cis.MbmAdj_AmendHistory", new { mbmNumber });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> MbmHistoryAsync(IDbSession session, string ic) =>
        session.RowsAsync("cis.Hdb_SearchCISMBMHistory", new { MBMICNo = ic });

    /// <inheritdoc />
    public Task<bool> OpenCaseAsync(IDbSession session, string mbmNumber) => session.ScalarAsync<bool>("cis.MbmAdj_OpenCase", new { mbmNumber });

    /// <inheritdoc />
    public Task EndorseAsync(IDbSession session, string mbmNumber, string endtType, string status, DateTime bordxDate, DateTime endtEffDate, string userCode, DateTime today) =>
        session.ExecAsync("cis.MbmAdj_Endorse", new { mbmNumber, endtType, status, bordxDate, endtEffDate, user = userCode, today });

    /// <inheritdoc />
    public Task CoverStatusAsync(IDbSession session, string mbmNumber, string status, string? covId, DateTime today) =>
        session.ExecAsync("cis.MbmAdj_CoverStatus", new { mbmNumber, status, covId, today });

    /// <inheritdoc />
    public Task RefundAsync(IDbSession session, string mbmNumber, string covId, decimal premEffective, decimal premReduction, decimal mcoEffective, decimal mcoRefund) =>
        session.ExecAsync("cis.MbmAdj_Refund", new { mbmNumber, covId, premEffective, premReduction, mcoEffective, mcoRefund });

    /// <inheritdoc />
    public Task SavePrincipalAsync(IDbSession session, string mbmNumber, AdjSaveRequest r, DateTime? effDate, DateTime? expDate) =>
        session.ExecAsync("cis.MbmAdj_SavePrincipal", new
        {
            mbmNumber, policyNo = (r.PolicyNo ?? "").Trim().ToUpperInvariant(), effDate, expDate, remarks = (r.CisRemarks ?? "").Trim(), cardType = (r.CardType ?? "").Trim(),
            mcoFee = r.McoFee, premium = r.Premium, dataStatus = (r.DataStatus ?? "A").Trim(), status = (r.Status ?? "A").Trim(), warranty = (r.Exclusion ?? "").Trim(),
            plnCode = (r.Plan ?? "").Trim(), email = (r.Email ?? "").Trim(), mobile = (r.Mobile ?? "").Trim(), bankName = (r.BankName ?? "").Trim(),
            bankAccount = (r.BankAccount ?? "").Trim(),
        });

    /// <inheritdoc />
    public Task SaveCoverAsync(IDbSession session, string mbmNumber, AdjCoverEdit c, DateTime? effDate, DateTime? expDate) =>
        session.ExecAsync("cis.MbmAdj_SaveCover", new
        {
            mbmNumber, covId = c.CoverId, remarks = (c.Remarks ?? "").Trim().ToUpperInvariant(), mco = c.Mco, premium = c.Premium,
            warranty = (c.Exclusion ?? "").Trim().ToUpperInvariant(), cardType = (c.CardType ?? "").Trim(), plnCode = (c.Plan ?? "").Trim(), effDate, expDate,
        });
}
