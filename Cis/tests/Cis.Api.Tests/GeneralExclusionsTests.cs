using Cis.Application.Files;
using Cis.Application.Membership;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Membership;

namespace Cis.Api.Tests;

public sealed class GeneralExclusionsServiceTests : IDisposable
{
    private sealed class Folder(string root) : IFileDrop
    {
        public Task<string?> SaveAsync(string folder, string fileName, byte[] content) => Task.FromResult<string?>(null);

        public string? FolderOf(string folder) => folder == FileFolders.PolicyWording ? root : null;
    }

    private readonly string _root = Directory.CreateTempSubdirectory("cis-pw-").FullName;

    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public async Task Policy_files_list_the_desktop_extensions_and_refuse_other_paths()
    {
        var company = Path.Combine(_root, "ZZ CO");
        Directory.CreateDirectory(company);
        foreach (var f in new[] { "a.PDF", "b.jpg", "c.exe", "d.docx" }) await File.WriteAllTextAsync(Path.Combine(company, f), "x");
        var service = new GeneralExclusionsService(new FakeSessions(), null!, new Folder(_root));
        Assert.Equal(["a.PDF", "b.jpg", "d.docx"], service.PolicyFiles("ZZ CO"));
        Assert.Empty(service.PolicyFiles("NO SUCH CO")!);
        Assert.Equal("a.PDF", (await service.PolicyFileAsync("ZZ CO", "a.pdf")).FileName);
        await Assert.ThrowsAsync<NotFoundException>(() => service.PolicyFileAsync("ZZ CO", "c.exe"));
        Assert.Throws<BusinessRuleException>(() => service.PolicyFiles(@"..\..\Windows"));
        Assert.Equal(GeneralExclusionsService.GroupFirst, Assert.Throws<BusinessRuleException>(() => service.PolicyFiles(" ")).Message);
        Assert.Null(new GeneralExclusionsService(new FakeSessions(), null!, new Folder("")).PolicyFiles("ZZ CO"));
    }
}

/// <summary>The enquiry against the real CISDB (skipped without CIS_TEST_DB).</summary>
[Collection(DatabaseCollection.Name)]
public class GeneralExclusionsDatabaseTests
{
    [Fact]
    public async Task A_company_shows_its_flags_and_remarks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new GeneralExclusionsService(TestDb.Sessions(), new GeneralExclusionsRepository(), new NoFolders());
        Assert.Contains("1MILLENNIUM INTERNATIONAL SDN BHD", await service.GroupCompaniesAsync());
        var v = await service.SearchAsync("1MILLENNIUM INTERNATIONAL SDN BHD");
        Assert.Equal(24, v.Flags.Count);
        Assert.True(v.Flags["Dental"]);
        Assert.False(v.Flags["Dermatitis"]);
        Assert.StartsWith("PAYOR : ST", v.Remarks);
        Assert.All((await service.SearchAsync("ZZ NO COMPANY")).Flags.Values, Assert.False);
    }

    private sealed class NoFolders : IFileDrop
    {
        public Task<string?> SaveAsync(string folder, string fileName, byte[] content) => Task.FromResult<string?>(null);

        public string? FolderOf(string folder) => null;
    }
}
