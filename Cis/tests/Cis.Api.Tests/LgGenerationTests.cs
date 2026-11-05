using Cis.Application.Files;
using Cis.Application.Invoice;
using Cis.Application.Lg;
using Cis.Domain.Errors;
using Cis.Infrastructure.Invoice;
using Cis.Infrastructure.Lg;
using Cis.Infrastructure.Membership;

namespace Cis.Api.Tests;

public class LgRulesTests
{
    private static Dictionary<string, object?> Row(string ann, string specAnn, string topup = "0.00") => new()
    {
        ["OriAnnLimit"] = ann, ["OriSpecAnnLimit"] = specAnn, ["TopupAmt"] = topup, ["SuppLimit"] = "500", ["LGIssuedInd"] = "Y", ["PLNCode"] = "P1",
        ["SNoLimitDedTotal"] = "N/A", ["SpecSuppLimitStatus"] = "",
    };

    private static InvClaim Claim(decimal bordx, string type = "G", string cover = "00", decimal others = 0) =>
        new() { ClmBordxAmt = bordx, ClnType = type, CoverId = cover, ClmOthersNoDed = others, PlnGpSpShared = "N" };

    [Fact]
    public void Gp_only_limits_sum_every_claim_less_the_no_deduction_amounts()
    {
        var l = LgRules.Limits(new LgLimitInput { Row = Row("1000", "N/A", "100"), LimitStatus = "A", SpecLimit = "N/A" },
            [Claim(200), Claim(150, "S", others: 50)]);
        Assert.Equal("1,000.00", l.AnnualLimit);
        Assert.Equal("300.00", l.Accumulated);
        Assert.Equal("800.00", l.Balance);
        Assert.Equal(("N/A", "N/A", "N/A"), (l.SpecAnnual, l.SpecAccumulated, l.SpecBalance));
        Assert.Equal("50.00", l.NoAnnLimitDed);
    }

    [Fact]
    public void A_specialist_limit_takes_the_specialist_clinic_claims_of_the_right_covers()
    {
        var input = new LgLimitInput { CoverId = "01", Row = Row("FULL", "2000"), LimitStatus = "C", SpecLimit = "2000", SpTopup = 100 };
        var l = LgRules.Limits(input, [Claim(100), Claim(300, "S", "01"), Claim(999, "S", "02")]);
        Assert.Equal(("FULL", "100.00", "FULL"), (l.AnnualLimit, l.Accumulated, l.Balance));
        Assert.Equal(("2,000.00", "300.00", "1,800.00"), (l.SpecAnnual, l.SpecAccumulated, l.SpecBalance));
        Assert.Equal("Y", l.LgIssuedInd);
        Assert.Equal(3, LgRules.ClaimMode("C", "01", false, true, ""));
        Assert.Equal(0, LgRules.ClaimMode("A", "01", false, true, ""));
        Assert.Equal(3, LgRules.ClaimMode("A", "01", false, true, "C"));
        Assert.Null(LgRules.ClaimMode("X", "00", false, false, ""));
    }

    [Fact]
    public void Shared_plans_and_status_e_follow_the_desktop()
    {
        var shared = LgRules.Limits(new LgLimitInput { Row = Row("1000", "N/A"), LimitStatus = "A", PrinDepGpShared = true }, [Claim(10)]);
        Assert.Equal(("N/A", "N/A", "N/A"), (shared.AnnualLimit, shared.Accumulated, shared.Balance));
        var supp = LgRules.Limits(new LgLimitInput { CoverId = "02", Row = Row("1000", "N/A"), LimitStatus = "E" }, []);
        Assert.Equal(("500.00", "500.00"), (supp.AnnualLimit, supp.Balance));
        var zero = LgRules.Limits(new LgLimitInput { Row = Row("0", "N/A"), LimitStatus = "C" }, []);
        Assert.Equal("FULL", zero.Balance);
    }

    [Fact]
    public void Validate_lg_uses_the_balances()
    {
        Assert.True(LgRules.IsValidLg("Y", "Y", "N", "FULL", ""));
        Assert.False(LgRules.IsValidLg("Y", "Y", "N", "0.00", ""));
        Assert.False(LgRules.IsValidLg("Y", "N", "N", "100", "0.00"));
        Assert.True(LgRules.IsValidLg("Y", "N", "N", "0", "N/A"));
        Assert.True(LgRules.IsValidLg("N", "Y", "Y", "0", ""));
    }
}

/// <summary>Generate LG against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class LgGenerationDatabaseTests
{
    private sealed class NoFiles : IFileDrop
    {
        public Task<string?> SaveAsync(string folder, string fileName, byte[] content) => Task.FromResult<string?>(null);

        public string? FolderOf(string folder) => null;
    }

    [Fact]
    public async Task Search_member_and_lg_read_the_member_and_its_lgs()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new LgService(TestDb.Sessions(), new LgRepository(), new ClinicInvoiceRepository(), new MbmEnquiryRepository(), new NoFiles());
        var blank = await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(" "));
        Assert.Equal(LgService.KeyMember, blank.Message);
        var found = await service.SearchAsync("ubha0511637*05");
        Assert.Equal(["00", "01", "02", "03"], found.Rows.Select(r => r.Cover));
        Assert.Equal("Patricia Ooi Puay Hua", found.Rows[0].Name);
        Assert.Null(found.Warranty);

        var m = await service.MemberAsync("UBHA0511637*05", "00");
        Assert.Equal("A", m.Status);
        Assert.Equal(("FULL", "FULL"), (m.Limits.AnnualLimit, m.Limits.Balance));
        Assert.Equal("Y", m.Limits.LgIssuedInd);
        Assert.True(m.ValidLg);
        Assert.Contains(m.Lgs, l => LgService.T(l, "LGNo") == "EA12590411");

        var lg = await service.LoadAsync("EA12590411", "United Overseas Bank (Malaysia) Bhd", "");
        Assert.Equal("UBHA0511637*05-00", lg.Member);
        Assert.Equal("Normal GL", lg.GlType);
        Assert.Equal("SUNWAY MEDICAL CENTRE", lg.Template.ClinicName);
        Assert.Equal("HC-Standard Template_UOB.dot", lg.Template.TemplateName);
        Assert.False(lg.Pharmacy);

        var lookups = await service.LookupsAsync();
        Assert.Contains("Normal GL", lookups.GlTypes);
        var cases = await service.CasesAsync("UBHA0511637*05", "00", false);
        Assert.NotNull(cases);
    }
}
