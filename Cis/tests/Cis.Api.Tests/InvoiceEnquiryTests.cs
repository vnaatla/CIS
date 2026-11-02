using Cis.Application.Files;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Invoice;
using Cis.Infrastructure.Membership;

namespace Cis.Api.Tests;

public sealed class InvoiceEnquiryDocumentTests : IDisposable
{
    private sealed class Folder(string? root) : IFileDrop
    {
        public Task<string?> SaveAsync(string folder, string fileName, byte[] content) => Task.FromResult<string?>(null);

        public string? FolderOf(string folder) => folder == FileFolders.Outpatient ? root : null;
    }

    private readonly string _root = Directory.CreateTempSubdirectory("cis-op-").FullName;

    public void Dispose() => Directory.Delete(_root, true);

    private InvoiceEnquiryService Service(string? root) => new(new FakeSessions(), null!, null!, null!, new Folder(root));

    [Fact]
    public async Task Documents_are_listed_by_folder_and_cannot_leave_them()
    {
        Directory.CreateDirectory(Path.Combine(_root, "Guarantee Letter", "EA1"));
        await File.WriteAllTextAsync(Path.Combine(_root, "Guarantee Letter", "EA1", "lg.pdf"), "x");
        await File.WriteAllTextAsync(Path.Combine(_root, "Guarantee Letter", "EA1", "note.txt"), "x");
        Directory.CreateDirectory(Path.Combine(_root, "SP_ScannedDoc"));
        await File.WriteAllTextAsync(Path.Combine(_root, "SP_ScannedDoc", "EA1.pdf"), "y");
        var service = Service(_root);
        Assert.Equal([new CaseDocument("Guarantee Letter", "lg.pdf"), new CaseDocument("SP_ScannedDoc", "EA1.pdf")], service.Documents("ea1"));
        Assert.Equal("y"u8.ToArray(), (await service.DocumentAsync("EA1", "SP_ScannedDoc", "EA1.pdf")).Content);
        await Assert.ThrowsAsync<NotFoundException>(() => service.DocumentAsync("EA1", "Guarantee Letter", "note.txt"));
        Assert.Throws<NotFoundException>(() => service.Documents("..\\EA1"));
        Assert.Null(Service(null).Documents("EA1"));
    }

    [Fact]
    public void Stored_worksheet_maps_the_seven_tables()
    {
        var cats = new List<(string, string)> { ("101", "GP"), ("102", "GP2") };
        var t = new List<Dictionary<string, object?>?>
        {
            new() { ["CLMTotalActual"] = 40m }, new() { ["CLMGPFeeBasicMed"] = 40m }, new() { ["CLMGPFeeBasicMed"] = 25m }, new() { ["CLMGPFeeBasicMed"] = 40m },
            null, null, new() { ["CLMGPFeeBasicMed"] = -15m }, new() { ["CLMGPFeeBasicMed"] = 40m },
        };
        var w = InvoiceEnquiryService.Stored(cats, t);
        Assert.Equal((40m, 25m, 40m, -15m, 40m), (w.Rows[0].Actual!.Value, w.Rows[0].ContractedRate!.Value, w.Rows[0].PayableToClinic!.Value,
            w.Rows[0].NotPayableToClinic!.Value, w.Rows[0].Approved!.Value));
        Assert.Null(w.Rows[1].Actual);
        Assert.Equal(40m, w.Claim!["CLMTotalActual"]);
    }
}

/// <summary>Enquiry against the real databases (skipped without CIS_TEST_DB).</summary>
[Collection(DatabaseCollection.Name)]
public class InvoiceEnquiryDatabaseTests
{
    [Fact]
    public async Task A_case_with_its_member_worksheet_and_tabs()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var clock = new FixedClock(new DateTime(2026, 11, 2));
        var clinic = new ClinicInvoiceService(TestDb.Sessions(), new ClinicInvoiceRepository(), new MemberSuppRepository(), new MbmAdjustmentRepository(),
            new MbmEnquiryRepository(), new WorksheetRepository(), clock);
        var service = new InvoiceEnquiryService(TestDb.Sessions(), new InvoiceEnquiryRepository(), clinic, new WorksheetRepository(), new NoFiles());
        var c = await service.FindAsync("ea12586874");
        Assert.Equal("EA12586874", ((string)c.Case["CASNumber"]!).Trim());
        Assert.Equal(31, c.Worksheet.Rows.Count);
        Assert.NotNull(c.Worksheet.Claim);
        Assert.NotEmpty(c.Bordx);
        await Assert.ThrowsAsync<NotFoundException>(() => service.FindAsync("ZZNOCASE1"));
        Assert.NotEqual("EA12586874", await service.StepAsync("EA12586874", true));
        var (payments, receipts) = await service.PaymentsAsync("EA12586874");
        Assert.NotNull(payments);
        Assert.NotNull(receipts);
        Assert.NotNull((await service.MedRequestsAsync("EA12586874")).Requests);
    }

    private sealed class NoFiles : IFileDrop
    {
        public Task<string?> SaveAsync(string folder, string fileName, byte[] content) => Task.FromResult<string?>(null);

        public string? FolderOf(string folder) => null;
    }
}
