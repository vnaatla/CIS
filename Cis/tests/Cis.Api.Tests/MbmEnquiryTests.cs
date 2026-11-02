using Cis.Application.Membership;
using Cis.Domain.Errors;
using Cis.Infrastructure.Membership;
using Xunit.Abstractions;
using static Cis.Application.Membership.MbmEnquiryRules;

namespace Cis.Api.Tests;

public class MbmEnquiryRulesTests
{
    private static EnqFlags F(string key, string pdl = "0")
    {
        var k = key.Split('|');
        return new EnqFlags(k[0], k[1], k[2], k[3], k[4], k[5], pdl);
    }

    private static EnqClaimRow Row(decimal bordx, decimal m = 0, decimal o = 0, decimal t = 0, decimal u = 0, string cov = "00", string type = "GP",
        string? cln = "G", string? shared = "N", decimal d = 0, decimal ffs = 0) =>
        new() { ClmBordxAmt = bordx, ClmMaternity = m, ClmOthersNoDed = o, ClmTyphiadJabNoDed = t, Utilization = u, CoverId = cov, TreatmentType = type, ClnType = cln,
                PlnGpSpShared = shared, ClmDental = d, Ffs = ffs };

    [Fact]
    public void Pdl_is_the_desktop_string()
    {
        Assert.Equal("0", Pdl(false, 250m));
        Assert.Equal("0.00", Pdl(true, null));
        Assert.Equal("250", Pdl(true, 250.0000m));
        Assert.True(F("C||A|B||", "0.00").PdlBlank);
        Assert.False(F("C||A|B||", "0").PdlBlank);
        Assert.True(F("C||A|B||", "250").PdlPositive);
    }

    [Fact]
    public void Gp_chain_follows_showMBM()
    {
        var r = Row(100, m: 10, o: 5, t: 2, u: 20);
        Assert.Equal(85, GpAdd(F("C||A|A|C|"), true, r));          // m + o
        Assert.Equal(95, GpAdd(F("C||A|A|C|"), false, r));         // o
        Assert.Equal(63, GpAdd(F("C||B|B||"), true, r));           // m + t + o + u
        Assert.Equal(88, GpAdd(F("C||D|B||"), false, r));          // m + t either way
        Assert.Equal(83, GpAdd(F("C||A|B||", "0.00"), true, r));   // PdlBlank
        Assert.Equal(98, GpAdd(F("C||A|B||", "250"), false, r));   // pdl > 0 → t only
        Assert.Equal(83, GpAdd(F("C||A|B||", "0"), false, r));     // neither → final Else (m + t + o)
        Assert.Null(GpAdd(F("||B|A|Y|"), true, r));                // only principal optical rows count
        Assert.Equal(83, GpAdd(F("||B|A|Y|"), true, Row(100, m: 10, o: 5, t: 2, type: "OPTICAL")));
        Assert.Equal(83, GpAdd(F("Y|C|C|D|Y|"), false, Row(100, m: 10, o: 5, t: 2, cov: "01")));  // cover <> 00 → final Else
    }

    [Fact]
    public void Specialist_branch_counts_gp_rows_only_unless_shared()
    {
        Assert.Null(GpAddWithSpec(true, true, Row(100, cln: "S")));
        Assert.Null(GpAddWithSpec(true, true, Row(100, cln: null)));
        Assert.Equal(100, GpAddWithSpec(false, false, Row(100, cln: "S", shared: "Y")));
        Assert.Equal(73, GpAddWithSpec(true, true, Row(100, m: 10, o: 5, t: 2, d: 10)));
        Assert.Equal(93, GpAddWithSpec(false, false, Row(100, m: 10, o: 5, t: 2, d: 10)));
    }

    [Fact]
    public void Vision_and_dental_labels()
    {
        Assert.Equal(new EnqBenefit(NA, "12.50", NA), Vision(F("C||A|A|C|"), 12.5m, "500"));
        Assert.Equal(new EnqBenefit("500", "12.50", "487.50"), Vision(F("|||F|Y|"), 12.5m, "500"));
        Assert.Equal(new EnqBenefit(NA, NA, NA), Vision(F("Y||B|B||"), 12.5m, "500"));
        Assert.Equal(1, DentalSource(F("C||A|A|C|")));
        Assert.Equal(2, DentalSource(F("C||B|B||")));
        Assert.Equal(0, DentalSource(F("||||||")));
        var rows = new List<EnqDentalRow>
        {
            new() { CoverId = "00", TreatmentType = "DENTAL", ClmBordxAmt = 40, BordxDentalAmount = 40, Tp1 = 1 },
            new() { CoverId = "01", TreatmentType = "DENTAL", ClmBordxAmt = 30, BordxDentalAmount = 30, Tp1 = 2 },
        };
        Assert.Equal((40m, 40m), PrincipalDental(F("C||A|A|C|"), rows));     // principal's dental rows only
        Assert.Equal((30m, 70m), PrincipalDental(F("C||A|B||"), rows));       // second rule needs cover 00, fourth-from-last does not
        var one = OneLimitDental(F("C||B|B||", "500"), "3", rows);
        Assert.Equal((40m, 0m, 30m, 3m), one);
        Assert.Equal(new EnqBenefit("500.00", "40.00", "497.00"), DentalLabels(F("C||B|B||", "500"), "3", (0, 0), one, 0));
        Assert.Equal(new EnqBenefit("500.00", "40.00", "460.00"), DentalLabels(F("C||B|B||", "500"), "", (0, 0), one, 0));
    }

    [Fact]
    public void Cover_rules_follow_showMBMCov()
    {
        Assert.Equal(2, CoverMode("B", "02", false));
        Assert.Null(CoverMode("B", "02", true));
        Assert.Equal(3, CoverMode("F", "02", false));
        Assert.Null(CoverMode("D", "02", false));
        Assert.Null(CoverMode("A", "01", false));
        var rows = new[] { Row(100, m: 10, o: 5, u: 20, cov: "01", ffs: 3), Row(50, cov: "01", cln: "S", ffs: 2) };
        Assert.Equal((0m, 0m), CoverGp(F("C||B|B||"), "A", rows));
        Assert.Equal((5m, 135m), CoverGp(F("Y||B|B||"), "C", rows));            // bordx - others - maternity (specialist rows too)
        Assert.Equal((5m, 115m), CoverGp(F("C||B|B||"), "C", rows));            // minus Utilization in this scenario
        Assert.Equal((0m, -20m), CoverGp(F("C||B|B||"), "F", rows));            // F rows add nothing, the deduction still applies
        Assert.Equal((3m, 95m), CoverGpWithSpec("C", rows));
        Assert.Equal((2m, 50m), CoverSpec("C", "C", "01", rows));
        Assert.Equal((0m, 0m), CoverSpec("E", "C", "01", rows));
        Assert.Equal((2m, 50m), CoverSpec("E", "", "01", rows));                // the desktop tests the GP status here
        var dental = new List<EnqDentalRow>
        {
            new() { CoverId = "01", TreatmentType = "DENTAL", ClmBordxAmt = 40 },
            new() { CoverId = "01", TreatmentType = "OPTICAL", ClmBordxAmt = 25 },
        };
        Assert.Equal((40m, 40m), SuppDental(F("C||B|B||"), dental));
        Assert.Equal(25m, SuppOptical(F("C||A|A|C|"), dental));
        Assert.Equal(65m, SuppOptical(F("Y|C|E|A|Y|C"), dental));               // every row in this scenario
        Assert.Equal(new EnqBenefit("300.00", "40.00", "235.00"), SuppDentalLabels(F("Y|Y|D|B|Y|"), "01", 300, (40, 40), 25));
        Assert.Equal(new EnqBenefit(NA, NA, NA), SuppDentalLabels(F("Y|Y|D|B|Y|"), "02", 300, (40, 40), 25));
    }

    [Fact]
    public void Spec_limit_and_utilisation_filter_follow_the_desktop()
    {
        var w = new List<string>();
        Assert.Equal("FULL", MbmEnquiryService.SpecLimit(new EnqSpecLimit(["F"], true, 100, "A"), w));
        Assert.Equal(NA, MbmEnquiryService.SpecLimit(new EnqSpecLimit([], false, null, null), w));
        Assert.Equal("0", MbmEnquiryService.SpecLimit(new EnqSpecLimit(["Y"], false, null, null), w));
        Assert.Equal("1500", MbmEnquiryService.SpecLimit(new EnqSpecLimit(["Y"], true, 1500, "A"), w));
        Assert.Empty(w);
        Assert.Equal("FULL", MbmEnquiryService.SpecLimit(new EnqSpecLimit(["Y"], true, 1500, ""), w));
        Assert.Equal([MbmEnquiryService.SpecWarning], w);
        Assert.Equal(0, MbmEnquiryService.Mode("A", false, false));
        Assert.Equal(1, MbmEnquiryService.Mode("F", false, false));
        Assert.Null(MbmEnquiryService.Mode("F", false, true));
        Assert.Equal(0, MbmEnquiryService.Mode("C", true, true));
    }

    [Fact]
    public async Task Member_must_have_eight_characters()
    {
        var service = new MbmEnquiryService(new FakeSessions(), null!, null!, null!);
        Assert.Equal(MbmEnquiryService.EnterMember, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.FindAsync("DCHC908"))).Message);
        Assert.Equal(MbmEnquiryService.EnterMember, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.CasesAsync(" "))).Message);
    }
}

/// <summary>MBM Enquiry against the real databases (skipped without CIS_TEST_DB).</summary>
[Collection(DatabaseCollection.Name)]
public class MbmEnquiryDatabaseTests(ITestOutputHelper output)
{
    [Fact]
    public async Task Find_and_cases()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new MbmEnquiryService(TestDb.Sessions(), new MbmEnquiryRepository(), new MbmAdjustmentRepository(), null!);
        var m = await service.FindAsync("DCHC9086608");
        output.WriteLine(System.Text.Json.JsonSerializer.Serialize(m with { Covers = [] }));
        Assert.Equal(("NUR RABIZAH BINTI ADENI", "H", "S079", "A"), (m.Name, m.InsType, m.Plan, m.LimitStatus));
        Assert.Equal(3, m.Covers.Count);
        output.WriteLine(System.Text.Json.JsonSerializer.Serialize(m.Covers));
        // Top-ups (negative here) are part of the balance: 7,500 - 4,371 = the member's stored 3,129 balance.
        Assert.Equal(("-4,371.00", "7,500.00", "3,129.00", "FULL"), (m.TopupAmt, m.AnnualLimit, m.BalanceLimit, m.VisitLimit));
        Assert.Equal(new EnqBenefit("250.00", "0.00", "250.00"), m.Dental);       // one-limit dental scenario (C||B|B||)
        Assert.Equal(new EnqBenefit("3,000.00", "0.00", "3,000.00"), m.Maternity);
        Assert.Equal(new EnqBenefit(NA, NA, NA), m.Specialist);

        // All plan flags blank (final scenario): utilised = bordereaux - (maternity + typhoid + others) of the 10 claims.
        var u = await service.FindAsync("DCHC4245383");
        output.WriteLine(System.Text.Json.JsonSerializer.Serialize(u with { Covers = [] }));
        Assert.Equal(("1,985.000", "18.50", "3,000.00", "1,015.00", "N/A"), (u.AccumAmt, u.FfsAmt, u.AnnualLimit, u.BalanceLimit, u.TopupAmt));
        Assert.StartsWith("CO-PAYMENT", u.SpecialRemark);
        await Assert.ThrowsAsync<NotFoundException>(() => service.FindAsync("ZZZZ99999999"));
        // Limit status B (shared by the supplementaries): every cover shows the covers' claims; only the first gets the limit and balance.
        var b = await service.FindAsync("DCHC8922156");
        Assert.Equal(("B", "908.000", "FULL"), (b.LimitStatus, b.AccumAmt, b.BalanceLimit));
        Assert.Equal(new EnqBenefit("5,000.00", "3,901.52", "1,098.48"), b.Covers[0].Gp);
        Assert.Equal("52.30", b.Covers[0].Ffs);
        Assert.All(b.Covers.Skip(1), c => Assert.Equal(new EnqBenefit("0.00", "3,901.52", "0.00"), c.Gp));
        var cases = await service.CasesAsync("DCHC4245383");
        Assert.Equal(10, cases.SpVisits);
        Assert.Equal(cases.SpBill, cases.SpBordx);
    }
}
