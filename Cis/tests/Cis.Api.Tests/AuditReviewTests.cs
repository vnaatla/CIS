using Cis.Application.Audit;
using Cis.Application.Files;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Audit;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Invoice;
using Dapper;

namespace Cis.Api.Tests;

public sealed class AuditReviewDocumentTests : IDisposable
{
    private sealed class Folder(string? root) : IFileDrop
    {
        public Task<string?> SaveAsync(string folder, string fileName, byte[] content) => Task.FromResult<string?>(null);

        public string? FolderOf(string folder) => folder == FileFolders.Outpatient ? root : null;
    }

    private readonly string _root = Directory.CreateTempSubdirectory("cis-gl-").FullName;

    public void Dispose() => Directory.Delete(_root, true);

    private AuditReviewService Service(string? root) => new(new FakeSessions(), null!, null!, null!, new Folder(root));

    [Fact]
    public async Task Copy_lists_and_downloads_guarantee_letters()
    {
        var service = Service(_root);
        Assert.Empty(service.Documents("EA1")!);
        Assert.Equal(AuditReviewService.Copied, await service.UploadAsync("ea1", "C:\\temp\\letter.pdf", new MemoryStream("x"u8.ToArray())));
        await File.WriteAllTextAsync(Path.Combine(_root, "Guarantee Letter", "EA1", "skip.exe"), "y");
        Assert.Equal(["letter.pdf"], service.Documents("EA1"));
        Assert.Equal("x"u8.ToArray(), (await service.DocumentAsync("EA1", "letter.pdf")).Content);
        var noCase = await Assert.ThrowsAsync<BusinessRuleException>(() => service.UploadAsync("", "a.pdf", Stream.Null));
        Assert.Equal(AuditReviewService.CaseFirst, noCase.Message);
        Assert.Throws<NotFoundException>(() => service.Documents("..\\EA1"));
        Assert.Null(Service(null).Documents("EA1"));
    }
}

/// <summary>Audit review against the real databases (skipped without CIS_TEST_DB); the update is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class AuditReviewDatabaseTests
{
    [Fact]
    public async Task Load_check_and_a_rolled_back_update()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new AuditReviewRepository();
        var service = new AuditReviewService(TestDb.Sessions(), repo, new ClinicInvoiceRepository(), new InvoiceEnquiryRepository(), new NoFiles());
        var c = await service.LoadAsync("ea12586874");
        Assert.Equal("T301", InvoiceAdjustmentService.T(c.Case, "ClinicCode"));
        Assert.Equal(["202", "101"], c.Diagnoses.Select(d => d.Code));
        await Assert.ThrowsAsync<NotFoundException>(() => service.LoadAsync("ZZNOCASE1"));
        var noQuery = await Assert.ThrowsAsync<BusinessRuleException>(() => service.UpdateAsync("EA12586874", new AuditReviewRequest("", "", "", "", "X", "01/11/2026")));
        Assert.Equal("Please Enter Query! ", noQuery.Message);
        var badDate = await Assert.ThrowsAsync<BusinessRuleException>(() => service.UpdateAsync("EA12586874", new AuditReviewRequest("Q", "", "", "", "X", "31/02/2026")));
        Assert.Equal("Invalid date format!", badDate.Message);
        Assert.NotEqual("EA12586874", await service.StepAsync("EA12586874", true));

        await using var s = await TestDb.Sessions().BeginAsync();
        await repo.UpdateAsync(s, "EA12586874", new AuditReviewRequest("ZZ QUERY", "ANSWER", "10", "REMARK", "DEVADMIN", "01/11/2026"), new DateTime(2026, 11, 1));
        var row = await s.Sql().Connection.QuerySingleAsync<(string Query, string By, DateTime Date)>(
            "SELECT RTRIM(AuditQuery), RTRIM(AuditBy), AuditDate FROM dbo.CISCas WHERE CASNumber = 'EA12586874'", transaction: s.Sql().Transaction);
        Assert.Equal(("ZZ QUERY", "DEVADMIN", new DateTime(2026, 11, 1)), row);
    }

    private sealed class NoFiles : IFileDrop
    {
        public Task<string?> SaveAsync(string folder, string fileName, byte[] content) => Task.FromResult<string?>(null);

        public string? FolderOf(string folder) => null;
    }
}
