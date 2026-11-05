using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Lg;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Lg;

/// <summary><see cref="ILgRepository"/> on the cis.Lg_* procedures.</summary>
public sealed class LgRepository : ILgRepository
{
    /// <inheritdoc />
    public async Task<(List<Dictionary<string, object?>> Rows, string? Warranty)> SearchAsync(IDbSession session, string member)
    {
        using var grid = await session.MultipleAsync("cis.Lg_Search", new { member });
        var rows = RowMapper.ToDictionaries(await grid.ReadAsync());
        var warranty = (await grid.ReadAsync<string?>()).FirstOrDefault();
        return (rows, warranty);
    }

    /// <inheritdoc />
    public async Task<LgMemberData> MemberAsync(IDbSession session, string member, string cover, string ic)
    {
        using var grid = await session.MultipleAsync("cis.Lg_Member", new { member, cover, ic });
        var limits = RowMapper.ToDictionaries(await grid.ReadAsync()).FirstOrDefault();
        var validate = RowMapper.ToDictionaries(await grid.ReadAsync()).LastOrDefault();
        var spTopup = (await grid.ReadAsync<decimal?>()).FirstOrDefault() ?? 0;
        var status = (await grid.ReadAsync<string?>()).ToList();
        var plan = RowMapper.ToDictionaries(await grid.ReadAsync()).FirstOrDefault();
        var aso = RowMapper.ToDictionaries(await grid.ReadAsync()).LastOrDefault();
        return new LgMemberData
        {
            Limits = limits, Validate = validate, SpTopup = spTopup, LimitStatus = status.Count == 0 ? null : status[0] ?? "",
            Maternity = S(plan, "PLNMaternityInd") == "Y", Dental = S(plan, "PDentalStatus") == "Y",
            PrinDepGpShared = plan?.GetValueOrDefault("IsPrinDepGPShared") is { } shared && Convert.ToBoolean(shared, CultureInfo.InvariantCulture),
            Aso = aso is null ? null : S(aso, "IsIPOPCombined") == "Y"
                ? (D(aso, "MBMASOLimit"), D(aso, "MBMIPUtilized") + D(aso, "MBMOPUtilized"))
                : (D(aso, "MBMOPLimit"), D(aso, "MBMOPUtilized")),
        };
    }

    /// <inheritdoc />
    public async Task<decimal?> MerchantradeAsync(IDbSession session, string member)
    {
        var r = (await session.RowsAsync("cis.Lg_Merchantrade", new { member })).FirstOrDefault();
        return r is not null && S(r, "Company") == "MERCHANTRADE ASIA SDN BHD" && r.GetValueOrDefault("BalLimit") is { } b
            ? Convert.ToDecimal(b, CultureInfo.InvariantCulture) : null;
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ListAsync(IDbSession session, string member, string? cover) =>
        session.RowsAsync("cis.Lg_List", new { member, cover });

    /// <inheritdoc />
    public async Task<LgLoadData> LoadAsync(IDbSession session, string lgNo)
    {
        using var grid = await session.MultipleAsync("cis.Lg_Load", new { lgNo });
        var rows = RowMapper.ToDictionaries(await grid.ReadAsync());
        var pending = (await grid.ReadAsync<string?>()).FirstOrDefault()?.Trim() == "Y";
        var pharmacy = (await grid.ReadAsync<int>()).FirstOrDefault() > 0;
        var main = (await grid.ReadAsync<string?>()).FirstOrDefault()?.Trim();
        var offline = RowMapper.ToDictionaries(await grid.ReadAsync()).FirstOrDefault();
        return new LgLoadData(rows, pending, pharmacy, string.IsNullOrEmpty(main) ? null : main,
            offline is null ? null : (S(offline, "PatientName"), S(offline, "IcNo")));
    }

    /// <inheritdoc />
    public async Task<bool> MedRequestInProgressAsync(IDbSession session, string caseNo) =>
        (await session.RowsAsync("cis.Lg_MedRequestInProgress", new { caseNo })).FirstOrDefault()?.GetValueOrDefault("InProgress") is true;

    /// <inheritdoc />
    public async Task<LgTemplate> TemplateAsync(IDbSession session, string clinic, string company, string plan, string glType, string ic, string category, string caseNo)
    {
        using var grid = await session.MultipleAsync("cis.Lg_Template", new { clinic, company, plan, glType, ic, category, caseNo });
        var cln = RowMapper.ToDictionaries(await grid.ReadAsync()).FirstOrDefault();
        var t = RowMapper.ToDictionaries(await grid.ReadAsync()).FirstOrDefault();
        var format = S(cln, "CLNGLFormat");
        return new LgTemplate
        {
            GlFormat = format == "M" ? "M" : "H", ClinicName = S(cln, "CLNName"), ClinicAddress1 = S(cln, "CLNAdd1"), ClinicAddress2 = S(cln, "CLNAdd2"),
            ClinicState = S(cln, "CLNState"), ClinicPanel = S(cln, "CLNPanel"), TemplateName = S(t, "GLTemplate"), DependantTemplate = S(t, "DepTemplate"),
            DependantInd = S(t, "DepInd"), VipInd = S(t, "VIPInd"), Remarks = S(t, "TmpRemarks"),
        };
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> CasesAsync(IDbSession session, string member, string? cover, bool specialist) =>
        session.RowsAsync("cis.Lg_Cases", new { member, cover, specialist });

    /// <inheritdoc />
    public async Task<LgLookups> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.Lg_Lookups");
        var types = RowMapper.ToDictionaries(await grid.ReadAsync()).Select(r => S(r, "GLType")).ToList();
        var decline = RowMapper.ToDictionaries(await grid.ReadAsync()).Select(r => S(r, "DeclineReason")).ToList();
        var delay = (await grid.ReadAsync<string?>()).Select(x => (x ?? "").Trim()).ToList();
        return new LgLookups(types, decline, delay);
    }

    /// <inheritdoc />
    public async Task<List<(string Code, string Name)>> ClinicsAsync(IDbSession session, string company) =>
        (await session.RowsAsync("cis.Lg_Clinics", new { company })).Select(r => (S(r, "CLNCode"), S(r, "CLNName"))).ToList();

    private static string S(IReadOnlyDictionary<string, object?>? r, string k) =>
        r is not null && r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal D(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
