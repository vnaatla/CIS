using Cis.Application.Email;
using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Reports;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Email Notification (Rejection) rules (no database).</summary>
public class EmlRejRuleTests
{
    private static EmlRejTick Tick(string? email, string company) =>
        new("EA1", email, "OOI", "OOI", "7207", "UOB2021", "INV1", "UBHA0120613*05", "01/07/2021", "27/05/2021", "240", "duplicated", "", company);

    [Fact]
    public void Bsn_members_without_a_personal_e_mail_get_a_letter()
    {
        Assert.True(EmlRejService.IsLetter(Tick("", "BANK SIMPANAN NASIONAL")));
        Assert.True(EmlRejService.IsLetter(Tick("bsn.hr@bsn.com.my", "BANK SIMPANAN NASIONAL")));
        Assert.False(EmlRejService.IsLetter(Tick("ooi@uob.com.my", "BANK SIMPANAN NASIONAL")));
        Assert.False(EmlRejService.IsLetter(Tick("", "United Overseas Bank (Malaysia) Bhd")));
    }

    [Fact]
    public void The_text_follows_the_desktop()
    {
        var t = EmlRejService.Text(Tick("a@b.c", "X"), true);
        Assert.StartsWith("Dear Sir/Madam\r\n\r\nCLAIM DETAILS\r\nMember Name            :OOI\r\n", t);
        Assert.Contains("Incurred Amount        :RM 240.00\r\n\r\n", t);
        Assert.EndsWith("Thank you.\r\nRegards, \r\nMediExpress (Malaysia) Sdn Bhd", t);
        Assert.EndsWith("Thank you.", EmlRejService.Text(Tick("", "X"), false));
    }

    private sealed class Captured : IEmailSender
    {
        public List<EmailMessage> Sent { get; } = [];

        public Task SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
        {
            Sent.Add(message);
            return Task.CompletedTask;
        }
    }

    [Fact]
    public async Task Nothing_ticked_is_refused()
    {
        var service = new EmlRejService(null!, null!, new Captured(), TimeProvider.System);
        Assert.Equal(EmlPayService.NoSelection, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SendAsync(new EmlRejSendRequest([])))).Message);
    }
}

/// <summary>Email Notification (Rejection) against the real databases (skipped without CIS_TEST_DB); the mark is rolled back and nothing is e-mailed.</summary>
[Collection(DatabaseCollection.Name)]
public class EmlRejDatabaseTests
{
    [Fact]
    public async Task Search_and_rolled_back_mark()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new EmlRejService(TestDb.Sessions(), new EmlRejRepository(), null!, TimeProvider.System);
        var rows = await service.SearchAsync(new EmlRejRequest("01/07/2021", "02/07/2021", "EA12514425", "EA12514425", "N - New", null));
        Assert.Equal("EA12514425", ((string)rows[0]["CASNumber"]!).Trim());
        Assert.Equal(EmlRejService.NoFrom, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SearchAsync(new EmlRejRequest("", "01/07/2021", null, null, "N - New", null)))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        await new EmlRejRepository().MarkAsync(tx, "EA12514425");
        Assert.Equal("Y", await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT CASEmailRejStatus FROM CISDB.dbo.CISCas WHERE CASNumber = 'EA12514425'", transaction: tx.Sql().Transaction));
    }
}
