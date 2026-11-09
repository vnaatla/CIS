using Cis.Application.Payment;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Payment;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Pay Clinic Post against the real databases (skipped without CIS_TEST_DB); the posting flag is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PayClnPostDatabaseTests
{
    [Fact]
    public async Task Lines_checks_and_rolled_back_mark()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PayClnPostService(TestDb.Sessions(), new PayClnPostRepository(), null!, TimeProvider.System);
        Assert.Equal(PayClnPostService.FutureDate,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new PayClnPostRequest(null, null, null, "01/01/2999", null, "0 - No")))).Message);
        var r = await service.SearchAsync(new PayClnPostRequest(null, null, null, "01/08/2021", "03/08/2021", "0 - No"));
        Assert.Equal(58, r.Lines.Count);
        var debit = r.Lines.First(l => l.Reference == "CP100470829");
        Assert.Equal(("84C159", "D", 652.50m), (debit.Account, debit.DebitCredit, debit.Amount));
        Assert.Equal(r.Debit, r.Credit);

        await using var tx = await TestDb.Sessions().BeginAsync();
        await new PayClnPostRepository().MarkAsync(tx, "047-106929-SNABILAH-C159", "C159", "DEVADMIN");
        Assert.Equal("1", await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT TOP 1 CASPayCLNPost FROM CISDB.dbo.CISCas WHERE InvoiceNo = '047-106929-SNABILAH-C159' AND ClinicCode = 'C159'", transaction: tx.Sql().Transaction));
    }
}
