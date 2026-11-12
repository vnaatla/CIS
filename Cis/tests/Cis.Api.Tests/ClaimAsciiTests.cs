using System.Text;
using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Reports;

namespace Cis.Api.Tests;

/// <summary>Claim ASCII formatting rules (no database).</summary>
public class ClaimAsciiRuleTests
{
    [Fact]
    public void L_pads_text_left_and_numbers_right_and_keeps_the_leftmost_characters()
    {
        Assert.Equal("AB   ", ClaimAsciiService.L("AB", 5));
        Assert.Equal("   12", ClaimAsciiService.L("12", 5));
        Assert.Equal("ABC", ClaimAsciiService.L("ABCDEF", 3));
        Assert.Equal("34.2", ClaimAsciiService.Num(34.2000m));
    }

    [Fact]
    public void Bintulu_line_maps_the_relationship_and_ends_in_cr()
    {
        var row = new Dictionary<string, object?>
        {
            ["MBMEmployeeNo"] = "E1", ["PatientName"] = "ALI", ["Relation"] = "W", ["DateVisit"] = new DateTime(2021, 7, 2), ["VisitTime"] = "10:00",
            ["CLMTotalActual"] = 34.2m, ["CASTreatReason1"] = null, ["MCDays"] = 1, ["BordxDate"] = new DateTime(2021, 7, 5),
        };
        var r = ClaimAsciiService.Bintulu([row], "KIV1", "NP");
        Assert.Equal("KIV1(NP)72021", r.File.FileName);
        Assert.Equal("BPSB|E1|ALI|U|02/07/2021|10:00||M|0|0|0|0|0|0|0|34.2|199|1|02/07/2021||\r\r\n", Encoding.Latin1.GetString(r.File.Content));
    }
}

/// <summary>Utusan line layout (no database: the restored data has no Utusan claims and view 01A needs the missing FWHSHISDB database).</summary>
public class ClaimAsciiUtusanTests
{
    [Fact]
    public void A_dependant_line_of_type_10A_has_the_treatment_code_and_the_name()
    {
        var row = new Dictionary<string, object?>
        {
            ["MBMEmployeeNo"] = "E1", ["DateVisit"] = new DateTime(2021, 7, 2), ["CLMBordxAmt"] = 34.2m, ["ClinicCode"] = "B294", ["InvoiceNo"] = "INV1",
            ["InvoiceDate"] = new DateTime(2021, 7, 3), ["coverid"] = "01", ["MBMCRELCode"] = "W", ["MBMCName"] = "SITI", ["CASTreatReason1"] = "304",
        };
        var r = ClaimAsciiService.Utusan([(11, [row])], new DateTime(2021, 7, 9));
        Assert.Equal("20210709_BPOCLMTRH.KZULF", r.File.FileName);
        Assert.Equal((1, 34.2m), (r.Visits, r.Total));
        Assert.Equal("UMMB,E1,10A,02/07/2021,34.20,34.20,FINANCE,B294,,03,INV1, ,03/07/2021,,,Upload From MediExpress,I,34.20,SITI,,,,,,,,,,,,,\r\r\n",
            Encoding.Latin1.GetString(r.File.Content));
    }
}

/// <summary>Claim ASCII exports against the real databases (skipped without CIS_TEST_DB); OMNI's provider mark is rolled back by not being run here.</summary>
[Collection(DatabaseCollection.Name)]
public class ClaimAsciiDatabaseTests
{
    [Fact]
    public async Task Bintulu_toshiba_singer_tims_cimb_and_checks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new ClaimAsciiService(TestDb.Sessions(), new ClaimAsciiRepository(), TimeProvider.System);
        Assert.NotEmpty(await service.ClinicsAsync(AsciiFormat.Toshiba));
        var b = await service.GenerateAsync(AsciiFormat.Bintulu, new ClaimAsciiRequest("KIV03274", "KIV03274", "NP - X", null));
        Assert.Equal(78, b.Visits);
        var t = await service.GenerateAsync(AsciiFormat.Toshiba, new ClaimAsciiRequest("KIV03274", "KIV03274", null, null));
        Assert.Equal("KIV03274(M-MEDIX).txt", t.File.FileName);
        Assert.StartsWith("M-MEDIX KIV03274", Encoding.Latin1.GetString(t.File.Content));
        var s = await service.GenerateAsync(AsciiFormat.Singer, new ClaimAsciiRequest("KIV03274", "KIV03274", null, null));
        Assert.Equal(t.Visits, s.Visits);
        Assert.Equal(18, (await service.GenerateAsync(AsciiFormat.Tims, new ClaimAsciiRequest("KDC55481", null, null, null))).Visits);
        Assert.Equal(571, (await service.GenerateAsync(AsciiFormat.Cimb, new ClaimAsciiRequest("KIS01540", null, null, "C - Cashless"))).Visits);
        var p = await service.GenerateAsync(AsciiFormat.OmniProviders, new ClaimAsciiRequest(null, null, null, null));
        Assert.True(p.Visits > 100);
        Assert.StartsWith("01|TPA|", Encoding.Latin1.GetString(p.File.Content));
        Assert.Equal(ClaimAsciiService.NoTo,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.GenerateAsync(AsciiFormat.Singer, new ClaimAsciiRequest("KIV03274", "", null, null)))).Message);
        Assert.Equal(ClaimAsciiService.NoRecord,
            (await Assert.ThrowsAsync<NotFoundException>(() => service.GenerateAsync(AsciiFormat.Cimb, new ClaimAsciiRequest("ZZZ", null, null, "R")))).Message);
    }
}
