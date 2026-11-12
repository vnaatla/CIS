using System.Globalization;
using System.IO.Compression;
using System.Text;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Reports;

/// <summary>Claim ASCII export formats: Bintulu (frmClaimDetailsASCII), Toshiba, Singer, TIMS, CIMB, OMNI (frmClaimDataOMNI), Utusan, OMNI providers (frmMedProviderOMNI).</summary>
public enum AsciiFormat
{
#pragma warning disable CS1591
    Bintulu, Toshiba, Singer, Tims, Cimb, Omni, Utusan, OmniProviders,
#pragma warning restore CS1591
}

/// <summary>Generate: bordx no. range (TIMS, CIMB and OMNI use the "from" only), clinic (Bintulu, "code - name"), cash type (CIMB, C / R).</summary>
public sealed record ClaimAsciiRequest(string? BordxFrom, string? BordxTo, string? Clinic, string? CashType);

/// <summary>Lengths.</summary>
public sealed class ClaimAsciiRequestValidator : Common.RequestValidator<ClaimAsciiRequest>
{
    /// <summary>Defines the rules.</summary>
    public ClaimAsciiRequestValidator()
    {
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.Clinic).MaximumLength(300);
        RuleFor(r => r.CashType).MaximumLength(20);
    }
}

/// <summary>The generated file and the screen's visit count / total.</summary>
public sealed record ClaimAsciiResult(FileOutput File, int Visits, decimal Total);

/// <summary>Data access.</summary>
public interface IClaimAsciiRepository
{
#pragma warning disable CS1591
    Task<List<string>> ClinicsAsync(IDbSession session, bool bintulu);
    Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, AsciiFormat format, string? bordxFrom, string? bordxTo, string? clinic);
    Task<string?> DeductibleAsync(IDbSession session, string plan, string marital);
    Task<List<Dictionary<string, object?>>> OmniProvidersAsync(IDbSession session);
    Task OmniSentAsync(IDbSession session, string clinic);
    Task<List<Dictionary<string, object?>>> UtusanAsync(IDbSession session, int claimType, string bordxFrom, string bordxTo);
#pragma warning restore CS1591
}

/// <summary>Claim ASCII exports.</summary>
public interface IClaimAsciiService
{
    /// <summary>The clinic combo (Bintulu: group company B003's clinics; Toshiba: panel clinics).</summary>
    Task<List<string>> ClinicsAsync(AsciiFormat format);

    /// <summary>cmdGenerate_Click (after the confirmation).</summary>
    Task<ClaimAsciiResult> GenerateAsync(AsciiFormat format, ClaimAsciiRequest request);
}

/// <inheritdoc />
public sealed class ClaimAsciiService(IDbSessions sessions, IClaimAsciiRepository repository, TimeProvider clock) : IClaimAsciiService
{
    /// <summary>Desktop texts.</summary>
    public const string NoFrom = "Please Enter Bordx No Fr ", NoTo = "Please Enter Bordx No To ", NoClinic = "Please Select Clinic Name ", NoRecord = "No record found!";

    /// <summary>Not a desktop text: the desktop wrote nothing when no cash type was chosen.</summary>
    public const string NoCashType = "Please select the cash type.";

    private static readonly CultureInfo En = CultureInfo.InvariantCulture;

    /// <inheritdoc />
    public async Task<List<string>> ClinicsAsync(AsciiFormat format)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.ClinicsAsync(s, format == AsciiFormat.Bintulu);
    }

    /// <inheritdoc />
    public async Task<ClaimAsciiResult> GenerateAsync(AsciiFormat format, ClaimAsciiRequest r)
    {
        var from = (r.BordxFrom ?? "").Trim();
        var to = (r.BordxTo ?? "").Trim();
        var clinic = (r.Clinic ?? "").Trim();
        clinic = (clinic.Contains('-') ? clinic[..clinic.IndexOf('-')] : clinic).Trim();
        var cash = (r.CashType ?? "").Trim();
        if (format is AsciiFormat.Bintulu or AsciiFormat.Toshiba or AsciiFormat.Singer or AsciiFormat.Utusan)
        {
            if (from.Length == 0) throw new BusinessRuleException(NoFrom);
            // The desktop compared the "to" box with an empty date mask, so this check never fired.
            if (to.Length == 0) throw new BusinessRuleException(NoTo);
            if (format == AsciiFormat.Bintulu && clinic.Length == 0) throw new BusinessRuleException(NoClinic);
        }
        if (format == AsciiFormat.Cimb && cash is not ("C" or "R") && !cash.StartsWith('C') && !cash.StartsWith('R')) throw new BusinessRuleException(NoCashType);
        var now = clock.GetLocalNow().DateTime;
        await using var tx = await sessions.BeginAsync();
        if (format == AsciiFormat.Utusan)
        {
            var parts = new List<(int, List<Dictionary<string, object?>>)>();
            for (var c = 1; c <= 13; c++) parts.Add((c, await repository.UtusanAsync(tx, c, from, to)));
            await tx.CommitAsync();
            return Utusan(parts, now);
        }
        var rows = await repository.RowsAsync(tx, format, from, to, clinic);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        ClaimAsciiResult result;
        switch (format)
        {
            case AsciiFormat.Bintulu: result = Bintulu(rows, to, clinic); break;
            case AsciiFormat.Toshiba: result = Toshiba(rows, from, now); break;
            case AsciiFormat.Singer: result = Singer(rows); break;
            case AsciiFormat.Tims:
                var deductibles = new Dictionary<(string, string), string>();
                foreach (var x in rows)
                {
                    var key = (Raw(x, "MBMPLNDeductible"), Raw(x, "MBMMaritalStatus"));
                    if (!deductibles.ContainsKey(key)) deductibles[key] = await repository.DeductibleAsync(tx, key.Item1, key.Item2) ?? "N/A";
                }
                result = Tims(rows, deductibles, now);
                break;
            case AsciiFormat.Cimb: result = Cimb(rows, cash[..1], now); break;
            case AsciiFormat.OmniProviders: result = OmniProviders(rows, now); break;
            default:
                var providers = await repository.OmniProvidersAsync(tx);
                foreach (var p in providers) await repository.OmniSentAsync(tx, Raw(p, "CLNCode").Trim());
                result = Omni(rows, providers, now);
                break;
        }
        await tx.CommitAsync();
        return result;
    }

    /// <summary>frmClaimDetailsASCII: "|"-separated lines ending in CR (then the CR LF of Print).</summary>
    public static ClaimAsciiResult Bintulu(List<Dictionary<string, object?>> rows, string bordxTo, string clinic)
    {
        var sb = new StringBuilder();
        decimal total = 0;
        foreach (var x in rows)
        {
            var rel = T(x, "Relation");
            rel = rel switch { "P" => "E", "H" or "W" => "U", _ => Raw(x, "Relation") };
            var mc = IsNum(x, "MCDays") ? T(x, "MCDays") : "0";
            var f = new[]
            {
                "BPSB", T(x, "MBMEmployeeNo"), T(x, "PatientName"), rel.Trim(), T(x, "DateVisit"), T(x, "VisitTime"), "", "M", "0", "0", "0", "0", "0", "0", "0",
                IsNum(x, "CLMTotalActual") ? T(x, "CLMTotalActual") : "0", Or(T(x, "CASTreatReason1"), "199"), mc,
                decimal.TryParse(mc, NumberStyles.Any, En, out var d) && d > 0 ? T(x, "DateVisit") : "", "", "",
            };
            sb.Append(string.Join('|', f)).Append('\r').Append("\r\n");
            total += N(x, "CLMTotalActual");
        }
        var first = rows[0].GetValueOrDefault("BordxDate") as DateTime?;
        return new(Text(sb, $"{bordxTo}({clinic}){first?.Month}{first?.Year}"), rows.Count, total);
    }

    /// <summary>frmClaimASCIIToshiba: fixed-width lines.</summary>
    public static ClaimAsciiResult Toshiba(List<Dictionary<string, object?>> rows, string bordxFrom, DateTime now)
    {
        var sb = new StringBuilder();
        decimal total = 0;
        var posting = 1000;
        foreach (var x in rows)
        {
            var b = x.GetValueOrDefault("BordxDate") as DateTime?;
            var v = x.GetValueOrDefault("DateVisit") as DateTime?;
            string D(DateTime? t) => t is null ? L("", 2) + L("", 2) + L("", 4) : L(t.Value.ToString("dd", En), 2) + L(t.Value.ToString("MM", En), 2) + L(t.Value.ToString("yyyy", En), 4);
            var icd = Raw(x, "ToshibaICD");
            posting++;
            var line = L("M-MEDIX", 8) + L(Raw(x, "Bordxrefno"), 15) + D(b) + (Raw(x, "MBMEmployeeNo") + new string(' ', 7))[..7]
                + L(T(x, "coverid") != "00" ? Raw(x, "PatientName") : "", 60) + L("NO", 2) + D(v)
                + (icd.Length > 0 ? L(icd, 2) + L(Raw(x, "ReasonDesc"), 60) : L("52", 2) + L("Medication & Treatment", 60))
                + L(Raw(x, "MCDays"), 3) + L(N(x, "CLMBordxAmt").ToString("#,##0.00", En).Replace(".", ""), 9) + L("P", 1) + D(b) + D(b)
                + L(posting.ToString("000000", En), 6) + L(now.ToString("HHmmss", En), 6) + L("", 10) + L("000000", 6);
            sb.Append(line).Append("\r\n");
            total += N(x, "CLMBordxAmt");
        }
        return new(Text(sb, $"{bordxFrom}(M-MEDIX).txt"), rows.Count, total);
    }

    /// <summary>frmClaimASCIISinger: header (the file name), CSV lines ending in CR, trailer.</summary>
    public static ClaimAsciiResult Singer(List<Dictionary<string, object?>> rows)
    {
        var first = (DateTime)rows[0]["BordxDate"]!;
        var name = $"CC{first:yy}{first:MM}{first:dd}01";
        var sb = new StringBuilder().Append(name).Append("\r\n");
        decimal total = 0;
        static string Q(string v) => "\"" + v + "\"";
        foreach (var x in rows)
        {
            sb.Append(string.Join(',', "1", T(x, "BordxDate"), "01", Q(T(x, "MBMIcBcPp")), Q(T(x, "MBMEmployeeNo")), Q(T(x, "PatientName")), T(x, "DateVisit"),
                IsNum(x, "CLMTotalActual") ? T(x, "CLMTotalActual") : "0", Q(T(x, "ClinicCode")), Q(T(x, "CLNName")))).Append('\r').Append("\r\n");
            total += N(x, "CLMTotalActual");
        }
        sb.Append($"2,{T(rows[^1], "BordxDate")},01,,,,,{Num(total)}").Append("\r\n");
        return new(Text(sb, name), rows.Count, total);
    }

    /// <summary>frmClaimASCIITIMS: "~"-separated lines (fields joined with + drop their separator when empty, as the desktop).</summary>
    public static ClaimAsciiResult Tims(List<Dictionary<string, object?>> rows, IReadOnlyDictionary<(string, string), string> deductibles, DateTime now)
    {
        var sb = new StringBuilder();
        decimal total = 0;
        foreach (var x in rows)
        {
            var o = new StringBuilder();
            void A(string v) => o.Append(v).Append('~');
            void P(string key) { if (x.GetValueOrDefault(key) is not (null or DBNull)) A(Raw(x, key)); }
            var visit = (x.GetValueOrDefault("DateVisit") as DateTime?)?.ToString("yyyyMMdd", En) ?? "";
            var actual = N(x, "CLMTotalActual");
            A("N"); A(Raw(x, "CASNumber")); A(""); A(visit); A(visit); A(Raw(x, "INVPayTo")); A(T(x, "ClinicCode")); A(T(x, "CLNName"));
            A(Raw(x, "CASTreatReason5") != "" && Raw(x, "CASDiagnosis5") != "" ? Clean(Raw(x, "CASDiagnosis5")) : Raw(x, "ReasonDesc"));
            A(Raw(x, "MedInjType1")); A(""); A(Num(actual)); A(Num(actual)); A("0");
            var fee = N(x, "GPFee") != 0 ? N(x, "GPFee") : N(x, "SPFee");
            A(Num(fee));
            var med = N(x, "GPMed") != 0 ? N(x, "GPMed") : N(x, "SPMed");
            A(Num(med)); A("0");
            A(actual > fee + med ? Num(actual - (fee + med)) : "0"); A("0");
            A(Raw(x, "MBMTINumber")); A(Raw(x, "MBMTICovID")); A(Raw(x, "Bordxrefno")); A(T(x, "BordxDate"));
            for (var i = 2; i <= 7; i++) A(Raw(x, "MedInjType" + i));
            for (var i = 2; i <= 5; i++) A(T(x, "ReasonDesc" + i).Length > 0 ? Raw(x, "ReasonDesc" + i) : "");
            for (var i = 1; i <= 7; i++) P("MedInjDays" + i);
            for (var i = 1; i <= 7; i++) P("MedInjQty" + i);
            A("");
            A(T(x, "CLMBTBGAmt").Length > 0 ? T(x, "CLMBTBGAmt") : "0");
            // The desktop added a found Currency deductible to "~" (a type mismatch); here it is written as text.
            A(deductibles.GetValueOrDefault((Raw(x, "MBMPLNDeductible"), Raw(x, "MBMMaritalStatus")), "N/A"));
            A(Raw(x, "MBMPLNDeductible"));
            P("MBMMaritalStatus");
            o.Append(T(x, "CLNType") == "G" ? "G" : "S");
            sb.Append(o).Append("\r\n");
            total += actual;
        }
        return new(Text(sb, $"{now:yyyyMMdd}_OUTPATIENTCLAIM.txt"), rows.Count, total);
    }

    /// <summary>frmClaimASCIICIMB: fixed-width lines for reimbursement (R) or cashless (C), trailer = the total actual as 18 digits.</summary>
    public static ClaimAsciiResult Cimb(List<Dictionary<string, object?>> rows, string cash, DateTime now)
    {
        var sb = new StringBuilder();
        decimal total = 0;
        static string Dm(object? v) => v is DateTime d ? d.ToString("dd-MMM-yyyy", En) : "";
        foreach (var x in rows)
        {
            var cas = Raw(x, "CASNumber");
            var inv = T(x, "InvoiceNo") != "" ? Raw(x, "InvoiceNo") : (cas.Length > 2 ? cas.Substring(2, Math.Min(9, cas.Length - 2)) : "");
            var icd = T(x, "ICDCode").Length > 0 ? Raw(x, "ICDCode") : "R05";
            string line;
            if (cash == "R")
                line = L(Dm(x.GetValueOrDefault("InvoiceReceiptDate")), 11) + L(Raw(x, "MBMEmployeeNo"), 7) + L(Raw(x, "PatientName"), 50) + L(T(x, "CLMTotalActual"), 10)
                    + L(T(x, "CLMTotalActual"), 10) + L(inv, 30) + L(Raw(x, "NPClinic"), 50) + L("", 10) + L(Dm(x.GetValueOrDefault("DateVisit")), 11)
                    + L("CLINIC(NON-PANEL)", 50) + L(icd, 50) + L("SELF", 10) + L(cas, 25) + L(Raw(x, "Bordxrefno"), 25) + L(Dm(x.GetValueOrDefault("BordxDate")), 11)
                    + L("GP", 5) + L("5210303", 7) + L("0", 10) + L("0", 10) + L("001692205056", 50);
            else
            {
                var invDate = x.GetValueOrDefault("InvoiceDate") is DateTime d && d.Date != new DateTime(1900, 1, 1) ? Dm(d) : Dm(x.GetValueOrDefault("BordxDate"));
                line = L(Raw(x, "ClinicCode"), 8) + L(Raw(x, "CLNName"), 50) + L(inv, 15) + L(invDate, 11) + L(Raw(x, "MBMEmployeeNo"), 7) + L(Raw(x, "PatientName"), 50)
                    + L(T(x, "CLMTotalActual"), 20) + (x.GetValueOrDefault("DateVisit") is DateTime ? L(Dm(x["DateVisit"]), 11) : "") + L("OUTPATIENT", 50) + L(icd, 50)
                    + L("SELF", 10) + L(cas, 25) + L(Raw(x, "Bordxrefno"), 25) + L(Dm(x.GetValueOrDefault("BordxDate")), 11) + L("GP", 15) + L("5210301", 7)
                    + L("0", 10) + L("0", 10) + L("001692205056", 50);
            }
            sb.Append(line).Append("\r\n");
            total += N(x, "CLMTotalActual");
        }
        sb.Append(Math.Round(total, 0, MidpointRounding.AwayFromZero).ToString("000000000000000000", En)).Append("\r\n");
        return new(Text(sb, $"{now:yyyyMMdd}CIMBMEDIX - {(cash == "R" ? "Reimbursement" : "Cashless")}.txt"), rows.Count, total);
    }

    /// <summary>frmClaimDataOMNI: the claim file (active members with a cover id) and the provider file (panel clinics not yet sent), in one zip.</summary>
    public static ClaimAsciiResult Omni(List<Dictionary<string, object?>> rows, List<Dictionary<string, object?>> providers, DateTime now)
    {
        var day = now.ToString("yyyyMMdd", En);
        var sb = new StringBuilder($"01|TPA|{day}\r\n");
        var count = 0;
        decimal total = 0;
        foreach (var x in rows)
        {
            var self = T(x, "coverid") == "00";
            if (T(x, self ? "MBMCLNStatus" : "MBMCCLNStatus") != "A") continue;
            var cov = T(x, self ? "MBMPrevMemNo" : "MBMAXANumber");
            if (cov.Length == 0) continue;
            count++;
            var visit = x.GetValueOrDefault("DateVisit") as DateTime?;
            var v = visit?.ToString("yyyyMMdd", En) ?? "";
            var actual = N(x, "CLMTotalActual");
            var gp = N(x, "GPMed");
            var med = gp != 0 ? gp : N(x, "SPMed");
            var net = gp != 0 ? actual - gp : IsNum(x, "SPMed") ? actual - N(x, "SPMed") : actual;
            var f = new List<string>
            {
                "02", "N", T(x, "CLNType") == "G" ? "OPGP" : "OPSP", Raw(x, "CASNumber"), "", cov, v, v, Raw(x, "INVPayTo") == "M" ? "2" : "1", T(x, "ClinicCode"),
                T(x, "CLNName"), "{ZZZ999}",
                "{" + (Raw(x, "CASTreatReason5") != "" && Raw(x, "CASDiagnosis5") != "" ? Raw(x, "CASDiagnosis5") : Raw(x, "ReasonDesc")) + "}",
                "{" + Raw(x, "MedInjType1") + "}", "", Num(actual), Num(actual), "0", Num(med), Num(med), "0", Num(net), Num(net),
            };
            f.AddRange(Enumerable.Repeat("0", 11));
            f.AddRange([$"{day} 01:00:00", "", "N", ""]);
            var mc = N(x, "MCDays");
            if (mc == 0) f.AddRange(["N", "", "", ""]);
            else f.AddRange(["Y", v, v, (mc == 1 ? visit : visit?.AddDays((double)mc))?.ToString("yyyyMMdd", En) ?? ""]);
            sb.Append(string.Join('|', f)).Append("\r\n");
            total += actual;
        }
        sb.Append($"03|{count}|{Num(total)}|{Num(total)}|0\r\n");
        using var zip = new MemoryStream();
        using (var a = new ZipArchive(zip, ZipArchiveMode.Create, true))
        {
            Add(a, $"{day}_OUTPATIENTCLAIM.txt", sb.ToString());
            if (providers.Count > 0)
            {
                var p = new StringBuilder($"01|TPA|{day}\r\n");
                foreach (var c in providers)
                    p.Append(string.Join('|', "02", "N", T(c, "CLNCode"), T(c, "CLNName"), T(c, "CLNPhone"), T(c, "CLNAdd1"), T(c, "CLNAdd2"), T(c, "CLNAdd3"),
                        T(c, "CLNCity"), T(c, "CLNPostCode"), "", "MALAYSIA", "2", "N", "Y")).Append("\r\n");
                p.Append($"03|{providers.Count}\r\n");
                Add(a, $"{day}_OUTPATIENTPROVIDER.txt", p.ToString());
            }
        }
        return new(new FileOutput(zip.ToArray(), "application/zip", $"{day}_OMNI.zip"), count, total);
    }

    private static readonly string[] UtusanTypes = ["01A", "01B", "01C", "02A", "02B", "03A", "03B", "04A", "04B", "09", "10A", "11A", "11B"];

    /// <summary>frmClaimASCIIUtusan: the 13 claim-type views in order, CSV lines ending in CR. The desktop counted the visits of types 01A and 01C only and never
    /// added up the amount; here every line is counted and the bordx amounts are totalled. As the desktop, the file is written even when no view returns a row.</summary>
    public static ClaimAsciiResult Utusan(List<(int Case, List<Dictionary<string, object?>> Rows)> parts, DateTime now)
    {
        var sb = new StringBuilder();
        int visits = 0;
        decimal total = 0;
        foreach (var (c, rows) in parts)
            foreach (var x in rows)
            {
                var amt = T(x, "CLMBordxAmt").Length > 0 ? N(x, "CLMBordxAmt").ToString("0.00", En) : "";
                var dependant = T(x, "coverid") != "00";
                var reason = c == 11 ? (T(x, "CASTreatReason1") == "304" ? "03" : T(x, "CLNType") == "S" ? "02" : "01") : "";
                var o = new StringBuilder();
                o.Append(string.Join(',', "UMMB", T(x, "MBMEmployeeNo"), UtusanTypes[c - 1], T(x, "DateVisit"), amt, amt, "FINANCE", T(x, "ClinicCode"), "", reason))
                    .Append(',').Append(T(x, "InvoiceNo")).Append(", ");
                o.Append(string.Join(',', "", T(x, "InvoiceDate"), "", "", "Upload From MediExpress",
                    dependant ? T(x, "MBMCRELCode") switch { "W" => "I", "H" => "S", _ => "A" } : "", amt)).Append(',');
                if (dependant) o.Append(T(x, "MBMCName")).Append(',');
                o.Append(string.Concat(Enumerable.Repeat(",", 12))).Append('\r');
                sb.Append(o).Append("\r\n");
                visits++;
                total += N(x, "CLMBordxAmt");
            }
        return new(Text(sb, $"{now:yyyyMMdd}_BPOCLMTRH.KZULF"), visits, total);
    }

    /// <summary>frmMedProviderOMNI: the provider file of the panel clinics (header 01|TPA, a 02 line per clinic, trailer 03 with the count).</summary>
    public static ClaimAsciiResult OmniProviders(List<Dictionary<string, object?>> rows, DateTime now)
    {
        var day = now.ToString("yyyyMMdd", En);
        var p = new StringBuilder($"01|TPA|{day}\r\n");
        foreach (var c in rows)
            p.Append(string.Join('|', "02", "N", T(c, "CLNCode"), T(c, "CLNName"), T(c, "CLNPhone"), T(c, "CLNAdd1"), T(c, "CLNAdd2"), T(c, "CLNAdd3"), T(c, "CLNCity"),
                T(c, "CLNPostCode"), "", "MALAYSIA", "2", "N", "Y")).Append("\r\n");
        p.Append($"03|{rows.Count}\r\n");
        return new(Text(p, $"{day}_OUTPATIENTPROVIDER.txt"), rows.Count, 0);
    }

    private static void Add(ZipArchive a, string name, string text)
    {
        using var w = new StreamWriter(a.CreateEntry(name).Open(), Encoding.Latin1);
        w.Write(text);
    }

    /// <summary>Medix.bas l(): the leftmost <paramref name="n"/> characters, right-aligned for numbers (RSet) and left-aligned otherwise (LSet).</summary>
    public static string L(string v, int n)
    {
        var cut = v.Length > n ? v[..n] : v;
        return decimal.TryParse(v.Trim(), NumberStyles.Any, En, out _) ? cut.PadLeft(n) : cut.PadRight(n);
    }

    /// <summary>A number as VB writes it (no trailing zeros).</summary>
    public static string Num(decimal v) => v.ToString("0.############", En);

    private static FileOutput Text(StringBuilder sb, string name) => new(Encoding.Latin1.GetBytes(sb.ToString()), "text/plain", name);

    /// <summary>The value as VB concatenates it (dates dd/MM/yyyy, numbers without trailing zeros, text untrimmed).</summary>
    private static string Raw(IReadOnlyDictionary<string, object?> r, string k) => r.GetValueOrDefault(k) switch
    {
        null or DBNull => "",
        DateTime d => d.TimeOfDay == TimeSpan.Zero ? d.ToString("dd/MM/yyyy", En) : d.ToString("dd/MM/yyyy h:mm:ss tt", En),
        decimal m => Num(m),
        double f => Num((decimal)f),
        var o => Convert.ToString(o, En) ?? "",
    };

    private static string T(IReadOnlyDictionary<string, object?> r, string k) => Raw(r, k).Trim();

    private static string Or(string v, string fallback) => v.Length > 0 ? v : fallback;

    private static bool IsNum(IReadOnlyDictionary<string, object?> r, string k) => decimal.TryParse(T(r, k), NumberStyles.Any, En, out _);

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) => decimal.TryParse(T(r, k), NumberStyles.Any, En, out var d) ? d : 0;

    /// <summary>Excel CLEAN: drops the control characters.</summary>
    private static string Clean(string v) => new(v.Where(c => c >= 32).ToArray());
}
