using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Lg;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Lg;

/// <summary><see cref="ILgSaveRepository"/> on cis.Lg_Checks / Lg_Issue / Lg_Decline / Lg_EmailRecipients.</summary>
public sealed class LgSaveRepository : ILgSaveRepository
{
    /// <inheritdoc />
    public async Task<LgChecks> ChecksAsync(IDbSession session, string lgNo, string member, string cover, string clinicCode, string clinicText)
    {
        using var grid = await session.MultipleAsync("cis.Lg_Checks", new { lgNo, member, cover, clinicCode, clinicText });
        var claim = RowMapper.ToDictionaries(await grid.ReadAsync()).FirstOrDefault();
        var payments = (await grid.ReadAsync<int>()).FirstOrDefault();
        var sp = (await grid.ReadAsync<string?>()).FirstOrDefault()?.Trim();
        var panel = (await grid.ReadAsync<string?>()).ToList();
        var expiry = (await grid.ReadAsync<DateTime?>()).FirstOrDefault();
        var pending = (await grid.ReadAsync<string?>()).FirstOrDefault()?.Trim() == "Y";
        return new LgChecks
        {
            HasClaim = claim is not null, Bordx = claim?.GetValueOrDefault("Bordx") is true, BtbgBordx = claim?.GetValueOrDefault("BtbgBordx") is true,
            InvoiceNo = Convert.ToString(claim?.GetValueOrDefault("InvoiceNo"), CultureInfo.InvariantCulture)?.Trim() ?? "", Paid = payments > 0,
            SpClinicOk = sp is "1" or "True" or "true", ClinicPanel = panel.Count == 0 ? null : (panel[0] ?? "").Trim(), Expiry = expiry, Pending = pending,
        };
    }

    /// <inheritdoc />
    public Task IssueAsync(IDbSession session, LgIssueSave s) =>
        session.ExecAsync("cis.Lg_Issue", new
        {
            lgNo = s.LgNo, user = s.User, treatmentDate = s.TreatmentDate, clinicCode = s.ClinicCode, remarks = s.Remarks, diagnosis = s.Diagnosis,
            lgAmount = s.LgAmount, glAmount = s.GlAmount, internalRemarks = s.InternalRemarks, doctor = s.Doctor, glType = s.GlType, maternity = s.Maternity,
            pending = s.Pending, delivery = s.Delivery, delayReason = s.DelayReason,
        });

    /// <inheritdoc />
    public Task DeclineAsync(IDbSession session, string lgNo, string doctor, string clinicCode, string reason, string user, string diagnosis, bool pharmacy) =>
        session.ExecAsync("cis.Lg_Decline", new { lgNo, doctor, clinicCode, reason, user, diagnosis, pharmacy });

    /// <inheritdoc />
    public async Task<(string To, string Bcc, string Cc, string Subsidiary)?> EmailRecipientsAsync(IDbSession session, string company, string member)
    {
        var r = (await session.RowsAsync("cis.Lg_EmailRecipients", new { company, member })).FirstOrDefault();
        string S(string k) => Convert.ToString(r?.GetValueOrDefault(k), CultureInfo.InvariantCulture)?.Trim() ?? "";
        return r is null ? null : (S("EmailAddress"), S("BCC"), S("CC"), S("Subsidiary"));
    }
}
