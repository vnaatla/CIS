using Cis.Application.Claim;
using Cis.Infrastructure.Claim;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using ClosedXML.Excel;
using Dapper;

namespace Cis.Api.Tests;

public class UpdateClaimsDataRulesTests
{
    private static MemoryStream Book(params object[][] rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(1, 1).Value = "header";
        for (var r = 0; r < rows.Length; r++)
            for (var c = 0; c < rows[r].Length; c++)
                if (rows[r][c] is { } v) ws.Cell(r + 2, c + 1).Value = XLCellValue.FromObject(v);
        var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        return ms;
    }

    private static object[] Row(int cols, params (int Col, object Value)[] values)
    {
        var row = new object[cols];
        foreach (var (col, value) in values) row[col - 1] = value;
        return row;
    }

    [Fact]
    public void Verification_follows_the_desktop_checks()
    {
        var service = new UpdateClaimsDataService(new FakeSessions(), null!, new SpreadsheetReader(), new FixedClock(new DateTime(2026, 11, 2)));
        var good = Row(10, (1, "N1234567"), (8, "ABCDEFGHIJKL"), (9, new DateTime(2026, 10, 1)), (10, 12.5));
        var bad = Row(10, (1, "X1234567"), (8, "SHORT"), (9, "notadate"), (10, "abc"));
        var r = service.Verify(ClaimDataKind.Ffs, Book(good, bad));
        Assert.Equal(UpdateClaimsDataService.Errors, r.Message);
        Assert.Equal(2, r.Rows);
        Assert.Equal(3, Assert.Single(r.Errors).Row);
        var ok = service.Verify(ClaimDataKind.Copay, Book(Row(20, (1, "x"), (2, "S123456789"), (3, "10.00"), (20, "31/10/2026"))));
        Assert.Equal(UpdateClaimsDataService.Verified, ok.Message);
        Assert.Empty(ok.Errors);
        Assert.Equal(UpdateClaimsDataService.Errors, service.Verify(ClaimDataKind.BordxDate, Book(Row(20, (1, "x"), (2, "S123456789")))).Message);
    }
}

/// <summary>Uploads against CISWorksheet (skipped without CIS_TEST_DB); rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class UpdateClaimsDataDatabaseTests
{
    [Fact]
    public async Task Each_upload_updates_its_column()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new UpdateClaimsDataRepository();
        await using var s = await TestDb.Sessions().BeginAsync();
        var today = new DateTime(2026, 11, 2);
        Assert.Equal(1, await repo.UpdateAsync(s, ClaimDataKind.Ffs, "EA12586874", 1.23m, null, today, "DEV"));
        Assert.Equal(1, await repo.UpdateAsync(s, ClaimDataKind.Copay, "EA12586874", 4.56m, null, today, "DEV"));
        Assert.Equal(1, await repo.UpdateAsync(s, ClaimDataKind.BordxAmt, "EA12586874", 7.89m, null, today, "DEV"));
        var row = await s.Sql().Connection.QuerySingleAsync<(decimal Ffs, decimal Copay, decimal Amt, string User)>(
            "SELECT CLMMCOProcFee, CLMCoPayColl, CLMBordxAmt, ClaimLastUpdateUser FROM CISWorksheet.dbo.Claim WHERE CASNumber = 'EA12586874'",
            transaction: s.Sql().Transaction);
        Assert.Equal((1.23m, 4.56m, 7.89m, "DEV"), row);
        var bordx = await s.Sql().Connection.QueryFirstOrDefaultAsync<string>(
            "SELECT TOP 1 BordxRefNo FROM CISWorksheet.dbo.Claim WHERE COALESCE(BordxRefNo, '') <> '' ORDER BY CASNumber DESC", transaction: s.Sql().Transaction);
        if (bordx is not null) Assert.True(await repo.UpdateAsync(s, ClaimDataKind.BordxDate, bordx, null, new DateTime(2026, 10, 31), today, "DEV") > 0);
    }
}
