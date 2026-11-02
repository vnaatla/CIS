using Cis.Application.Invoice;
using Cis.Infrastructure.Invoice;

namespace Cis.Api.Tests;

public class WorksheetEngineTests
{
    private static readonly string[] Codes =
        ["101", "102", "103", "104", "105", "106", "301", "302", "303", "304", "305", "306", "307", "308", "309", "401", "402", "403", "404", "405",
         "501", "502", "503", "504", "505", "601", "602", "603", "604", "701", "702"];

    private static WorksheetLookups Db(string type = "2", string red = "", decimal mco = 0) => new()
    {
        Categories = Codes.Select(c => (c, "BNF " + c)).ToList(),
        ClinicBenefitType = type, ClinicReductionStatus = red, McoProcFeePercent = mco,
        SpecialRates = new Dictionary<string, decimal?> { ["101"] = 25m, ["102"] = 35m },
        StandardRates = new Dictionary<string, decimal?> { ["101"] = 99m, ["301"] = 25m, ["302"] = 18m },
        PlanBenefitStatus = true,
        Benefits = [new("101", 50m, 30m, "1"), new("301", 100m, 10m, "2"), new("302", null, 7m, "4")],
        ReductionAmount = 1m, ReductionPercent = 5m,
    };

    private static WorksheetInput Input(params (int Row, decimal Amount)[] items)
    {
        var list = new decimal?[31];
        foreach (var (row, amount) in items) list[row] = amount;
        return new WorksheetInput
        {
            Caller = WorksheetCaller.ClinicInvoice, PayTo = "C", ClinicCode = "D118", ClinicBenefitCard = "Classic", Items = list, NonGpAmount = 5m,
            Row70BalanceLimit = "1,000.00", BalanceLimit = "1,000.00", VisitLimit = "FULL", AnnualIndicator = true, VisitIndicator = true,
            CoPayCollected = 2m, CoPay = 0m, AdjustedAnnualLimit = 100m, GpStatus = true, McoProcStatus = true,
        };
    }

    [Fact]
    public void Cashless_standard_rates_follow_the_desktop()
    {
        // 101 billed 40 (card rate 25), 301 billed 20 (standard 25), 302 billed 30 (standard 18).
        var r = WorksheetEngine.Calculate(Input((0, 40m), (6, 20m), (7, 30m)), Db(mco: 10m), new DateTime(2026, 11, 2), "devadmin");
        var r101 = r.Rows[0];
        Assert.Equal((40m, 25m, 40m, -15m, 20m, 40m), (r101.Actual!.Value, r101.ContractedRate!.Value, r101.PayableToClinic!.Value,
            r101.NotPayableToClinic!.Value, r101.BenefitLimit!.Value, r101.Approved!.Value));
        Assert.Null(r101.Excess);
        Assert.Equal(10m, r.Rows[6].BenefitLimit);          // GetMax returns Y (desktop quirk)
        Assert.Null(r.Rows[6].NotPayableToClinic);           // 20 <= 25
        Assert.Equal((7m, -12m), (r.Rows[7].BenefitLimit!.Value, r.Rows[7].NotPayableToClinic!.Value));
        Assert.Null(r.Rows[1].Actual);
        Assert.Equal((90m, 90m, 90m, -27m), (r.TotalActual, r.TotalApproved, r.TotalPayableToClinic, r.TotalNotPayableToClinic));
        // No reduction status: 90 - 2 = 88, MCO fee 10% = 8.8 → 79.2.
        Assert.Equal((8.8m, 79.2m), (r.McoProcFee, r.FinalPayableToClinic));
        // Visit FULL → visit limit = approved 90; balance 1,000 + 100; 90 <= 1,100 → 90 - 2 - 0.
        Assert.Equal((1100m, 88m), (r.FinalBalanceLimit, r.FinalApproved!.Value));
        Assert.Equal(-2m, r.FinalExcess);
        Assert.Equal(40m, r.RowValues[168]);
        Assert.Equal(40m, r.RowValues[211]);                 // the desktop stores the actual as the contracted rate
        Assert.Equal(-15m, r.RowValues[383]);
        Assert.Equal((95m, 88m, 1012m, 1100m, 8.8m, "DEVADMIN"), ((decimal)r.RowValues[160]!, (decimal)r.RowValues[161]!, (decimal)r.RowValues[163]!,
            (decimal)r.RowValues[164]!, (decimal)r.RowValues[468]!, (string)r.RowValues[151]!));
    }

    [Fact]
    public void Reduction_balance_limit_and_reimbursement()
    {
        var input = Input((0, 40m)) with { VisitLimit = "30", AnnualIndicator = false };
        var r = WorksheetEngine.Calculate(input, Db(red: "2"), new DateTime(2026, 11, 2), "dev");
        // Reduction 5% of 40 = 2; IIf(40 - 2 - 2 <> 0, 40 - 2, 0) - 2 = 36; MCO 0%.
        Assert.Equal(36m, r.FinalPayableToClinic);
        // ANN off → balance = approved 40; visit 30 <= 40 → 40 - 2.
        Assert.Equal(38m, r.FinalApproved);
        var big = WorksheetEngine.Calculate(input with { VisitLimit = "500", AnnualIndicator = true, Row70BalanceLimit = "10" }, Db(), DateTime.Today, "dev");
        // visit 500 > balance 10 + 100 = 110? no: 40 <= 110 → approved - copay collected.
        Assert.Equal(38m, big.FinalApproved);
        var capped = WorksheetEngine.Calculate(input with { VisitLimit = "500", AnnualIndicator = true, Row70BalanceLimit = "10", AdjustedAnnualLimit = 0 },
            Db(), DateTime.Today, "dev");
        Assert.Equal(8m, capped.FinalApproved);              // 40 > 10 → 10 - 2

        var m = WorksheetEngine.Calculate(input with { PayTo = "M", Caller = WorksheetCaller.MemberInvoice, ApprovedOverride = "12.50" }, Db(),
            DateTime.Today, "dev");
        Assert.Equal(40m, m.Rows[0].Excess);                 // reimbursement: due to member = approved
        Assert.Null(m.Rows[0].ContractedRate);
        Assert.Equal(-40m, m.Rows[0].NotPayableToClinic);    // no clinic rates for reimbursement
        Assert.Equal((12.5m, 0m), ((decimal)m.RowValues[161]!, (decimal)m.RowValues[162]!));
        Assert.False(m.RowValues.ContainsKey(468));
    }

    [Fact]
    public void Helpers_reproduce_GetMin_DueToFromMBM_and_DueToFromCLN()
    {
        Assert.Null(WorksheetEngine.Min(0, 5));
        Assert.Equal(3m, WorksheetEngine.Min(3, 5));
        Assert.Equal(-5m, WorksheetEngine.DueToFromMember("C", 10, 5));
        Assert.Null(WorksheetEngine.DueToFromMember("C", 5, 5));
        Assert.Equal(-2m, WorksheetEngine.DueToFromClinic(7, 5));
        Assert.Null(WorksheetEngine.DueToFromClinic(5, 7));
    }
}

/// <summary>Worksheet lookups against CISDB / HISDB (skipped without CIS_TEST_DB).</summary>
[Collection(DatabaseCollection.Name)]
public class WorksheetDatabaseTests
{
    [Fact]
    public async Task Lookups_read_the_clinic_rates_and_plan_benefits()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().OpenAsync();
        var db = await new WorksheetRepository().LookupsAsync(s, new WorksheetInput
        {
            ClinicCode = "D118", ClinicBenefitCard = "Classic", PlanCode = "A001", InsCode = "G", HltCode = "N",
        });
        Assert.Equal(31, db.Categories.Count);
        Assert.Equal(("2", 10m), (db.ClinicBenefitType, db.McoProcFeePercent));
        Assert.Equal(25m, db.StandardRates["301"]);
        Assert.Equal(35m, db.SpecialRates["102"]);
        Assert.NotEmpty(db.Benefits);
        Assert.Equal(1m, db.ReductionAmount);
    }
}
