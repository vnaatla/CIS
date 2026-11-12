using Cis.Application.Email;
using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Reports;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Email Notification (Payment) rules (no database).</summary>
public class EmlPayRuleTests
{
    private sealed class NoMail : IEmailSender
    {
        public Task SendAsync(EmailMessage message, CancellationToken cancellationToken = default) => throw new InvalidOperationException("no e-mail in tests");
    }

    [Fact]
    public void The_body_has_the_advice_and_the_claims_table_with_totals()
    {
        var d = new Dictionary<string, object?>
        {
            ["CasNumber"] = "EA1", ["DateVisit"] = new DateTime(2021, 6, 16), ["CLMTotalActual"] = 95m, ["CLMBordxAmt"] = 90.5m, ["PatientName"] = "ALI",
            ["MBMGroupCompany"] = "FMC",
        };
        var html = EmlPayService.Body(new EmlPayTick("CP1", "a@b.c", "ALI", "M1", "123", "20/06/2021"), [d], "05/08/2021");
        Assert.StartsWith("<HTML><BODY><font color=black face=\"Arial\" size=1>Dear Sir/Madam,", html);
        Assert.Contains("<tr><td>1</td> <td>EA1</td> <td>16/06/2021</td> <td>95.00</td> <td>90.50</td><td>ALI</td><td>FMC</td></tr>", html);
        Assert.Contains("Paid Amount             :       RM 90.50</p></p>", html);
        Assert.Contains("Date Banked-in         :       05/08/2021</p></p>", html);
    }

    [Fact]
    public async Task Send_checks_come_before_any_e_mail()
    {
        var service = new EmlPayService(null!, null!, new NoMail());
        Assert.Equal(EmlPayService.NoSendDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SendAsync(new EmlPaySendRequest("", null, [])))).Message);
        Assert.Equal(EmlPayService.NoSelection, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SendAsync(new EmlPaySendRequest("05/08/2021", null, [])))).Message);
        Assert.Equal(EmlPayService.NoEmail, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SendAsync(new EmlPaySendRequest("05/08/2021", null, [new("CP1", "", "A", "M", "1", "x")])))).Message);
    }
}

/// <summary>Email Notification (Payment) against the real databases (skipped without CIS_TEST_DB); the mark is rolled back and nothing is e-mailed.</summary>
[Collection(DatabaseCollection.Name)]
public class EmlPayDatabaseTests
{
    [Fact]
    public async Task Search_details_and_rolled_back_mark()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new EmlPayService(TestDb.Sessions(), new EmlPayRepository(), null!);
        var rows = await service.SearchAsync(new EmlPayRequest("03/08/2021", "03/08/2021", "CP100471624", "CP100471624", "N - New", null, 0));
        Assert.Equal("CP100471624", ((string)Assert.Single(rows)["PVNo"]!).Trim());
        Assert.Equal(EmlPayService.NoStatus, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SearchAsync(new EmlPayRequest("03/08/2021", "03/08/2021", null, null, "", null, 0)))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        Assert.NotEmpty(await new EmlPayRepository().DetailsAsync(tx, "CP100471624"));
        await new EmlPayRepository().MarkAsync(tx, "CP100471624");
        Assert.True(await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISPayment.dbo.Payment2MBM WHERE PVNo = 'CP100471624' AND PVEmailStatus = 'Y'", transaction: tx.Sql().Transaction) > 0);
    }
}
