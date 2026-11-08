using Cis.Application.Email;
using Cis.Application.Files;
using Cis.Application.Invoice;
using Cis.Application.Lg;
using Cis.Domain.Errors;
using Cis.Infrastructure.Invoice;
using Cis.Infrastructure.Lg;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Membership;
using Dapper;

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
        var service = Service();
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

    private static LgService Service(IEmailSender? mail = null) =>
        new(TestDb.Sessions(), new LgRepository(), new LgSaveRepository(), new ClinicInvoiceRepository(), new MbmEnquiryRepository(), new NoFiles(),
            new PdfLetterWriter(), mail ?? new NoMail(), TimeProvider.System);

    private sealed class NoMail : IEmailSender
    {
        public Task SendAsync(EmailMessage message, CancellationToken cancellationToken = default) => throw new NotConfiguredException("E-mail is not configured.");
    }

    [Fact]
    public async Task Add_edit_and_issue_follow_the_desktop_checks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = Service();
        // The member's policy expired (31/12/2021), so Add stops as the desktop does.
        var add = await service.AddCheckAsync("UBHA0511637*05", "00", 0);
        Assert.Equal(LgGateKind.Stop, add.Kind);
        Assert.Equal("Letter of Guarantee is Not Allowed !!", add.Message);
        // The restored member's policy ended in 2021, so Edit stops at the expiry check (the LG has no claim, payment or panel problem).
        var edit = await service.EditCheckAsync("EA12590411", "SUNWAY MEDICAL CENTRE[S427]", "", 0);
        Assert.Equal((LgGateKind.Stop, LgService.Expired), (edit.Kind, edit.Message));
        var noLg = await service.EditCheckAsync("", null, null, 0);
        Assert.Equal(LgService.SelectLg, noLg.Message);
        var issue = await service.IssueAsync(new LgIssueRequest { LgNo = "EA12590411", ClinicText = "SUNWAY MEDICAL CENTRE[S427]", GlAmount = "100" }, "DEVADMIN");
        Assert.Equal((LgGateKind.Stop, LgService.Expired), (issue.Kind, issue.Message));
        var decline = await service.DeclineAsync(new LgDeclineRequest { LgNo = "EA12590411" }, "DEVADMIN");
        Assert.Equal(LgService.DeclineFirst, decline.Message);
        var ask = await service.DeclineAsync(new LgDeclineRequest { LgNo = "EA12590411", Reason = "X" }, "DEVADMIN");
        Assert.Equal((LgGateKind.Confirm, "Are you sure you want to decline [EA12590411]?", true), (ask.Kind, ask.Message, ask.YesNo));
    }

    [Fact]
    public async Task Issue_and_decline_write_the_desktop_tables_rolled_back()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new LgSaveRepository();
        await using var s = await TestDb.Sessions().BeginAsync();
        await repo.IssueAsync(s, new LgIssueSave("EA12590411", "DEVADMIN", new DateTime(2021, 8, 28), "S427", "REMARK", "ZZ DIAG", 321, 321, "INT", "DR ZZ",
            "Normal GL", 0, false, false, "LATE"));
        var lg = await s.Sql().Connection.QuerySingleAsync<(decimal Amount, string Diag, string By)>(
            "SELECT GLAmt, Diagnosis, LGIssuedBy FROM dbo.LG WHERE LGNo = 'EA12590411'", transaction: s.Sql().Transaction);
        Assert.Equal((321m, "ZZ DIAG", "DEVADMIN"), lg);
        var progress = await s.Sql().Connection.QuerySingleAsync<string>("SELECT CaseProgress FROM dbo.tblRecord WHERE CaseNo = 'EA12590411'", transaction: s.Sql().Transaction);
        Assert.Equal("LG Issued", progress);
        await repo.DeclineAsync(s, "EA12590411", "DR ZZ", "S427", "ZZ REASON", "DEVADMIN", "ZZ DIAG", false);
        var declined = await s.Sql().Connection.QuerySingleAsync<string>("SELECT CaseProgress FROM dbo.tblRecord WHERE CaseNo = 'EA12590411'", transaction: s.Sql().Transaction);
        Assert.Equal("LG Declined", declined);
    }
}

public class LgSaveRulesTests
{
    private static LgLimits L(string bal, string spec) => new() { Balance = bal, SpecBalance = spec };

    [Fact]
    public void Balance_test_reproduces_cmdLG()
    {
        Assert.True(LgService.BalanceValid("100", L("FULL", "N/A"), "Y", "N", "Normal GL", 0));
        Assert.False(LgService.BalanceValid("600", L("500", "N/A"), "Y", "N", "Normal GL", 0));
        Assert.True(LgService.BalanceValid("400", L("500", "N/A"), "Y", "N", "Normal GL", 0));
        // GP / SP not shared with no specialist balance: the amount is compared with 0.
        Assert.False(LgService.BalanceValid("100", L("FULL", "N/A"), "N", "N", "Normal GL", 0));
        // With a previous GL amount the desktop compares IsNumeric (-1) with the balance.
        Assert.True(LgService.BalanceValid("99999", L("0", "N/A"), "N", "N", "Normal GL", 500));
        Assert.True(LgService.BalanceValid("99999", L("0", "0"), "N", "Y", "Normal GL", 0));
    }

    [Fact]
    public void Gates_follow_cmdAdd()
    {
        var full = new LgLimits { LgIssuedInd = "Y", UsingLimit = "A", LgIssuedAfterLimitExInd = "N", AnnualLimit = "FULL", Balance = "FULL", SpecBalance = "N/A" };
        var expiry = DateTime.Today.AddDays(30);
        var first = LgGates.Add(full, true, true, expiry, DateTime.Today, "X", 0);
        Assert.Equal((LgGateKind.Confirm, "Press OK to Add LG and Cancel to Abort."), (first.Kind, first.Message));
        Assert.Equal(LgGateKind.Ok, LgGates.Add(full, true, true, expiry, DateTime.Today, "X", 1).Kind);
        Assert.Equal("No LG For this Plan!!!!", LgGates.Add(full, false, true, expiry, DateTime.Today, "X", 0).Message);
        var exceeded = full with { Balance = "0.00", AnnualLimit = "100.00" };
        Assert.Equal("Limit Exceeded. Letter of Guarantee is Not Allowed !!", LgGates.Add(exceeded, true, true, expiry, DateTime.Today, "X", 0).Message);
        var max = LgGates.Add(full with { Balance = "250.00", AnnualLimit = "500.00" }, true, true, expiry, DateTime.Today, "X", 0);
        Assert.Equal("Member Entitled Max LG Amount RM 250.00. Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.", max.Message);
        Assert.Equal("Letter of Guarantee is Not Allowed !!", LgGates.Add(full, true, true, DateTime.Today.AddDays(-1), DateTime.Today, "X", 0).Message);
        Assert.Equal("S427", LgService.ClinicCode("SUNWAY MEDICAL CENTRE     [ S427]"));
    }
}
