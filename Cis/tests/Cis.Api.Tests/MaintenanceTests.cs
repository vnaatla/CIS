using System.Text;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.CodeTables;
using Cis.Application.Maintenance.ListReports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Maintenance;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

public class CodeTableRuleTests
{
    private static CodeTableDefinition T(string key) => CodeTableCatalog.Find(key)!;
    private static Dictionary<string, string?> V(params (string, string?)[] values) => values.ToDictionary(v => v.Item1, v => v.Item2, StringComparer.OrdinalIgnoreCase);

    [Theory]
    [InlineData("", "Name", "10", "Please Enter Category Code")]
    [InlineData("AB", "Name", "10", "Category Code must be 3 integer!")]
    [InlineData("ABC", "", "10", "Please Enter Category Description")]
    [InlineData("ABC", "Name", "", "Please Enter Amount")]
    [InlineData("ABC", "Name", "x1", "Please Enter Amount")]
    public void Check_item_save_reports_the_desktop_messages_in_order(string code, string name, string amount, string message)
    {
        var e = Assert.Throws<BusinessRuleException>(() => CodeTableService.Validate(T("check-items"), V(("CatCode", code), ("CatName", name), ("Amount", amount)), false));
        Assert.Equal(message, e.Message);
    }

    [Fact]
    public void Drug_value_is_stored_formatted_like_the_desktop_and_optional_fields_may_be_empty()
    {
        var values = CodeTableService.Validate(T("drugs"), V(("DrugCode", "D01"), ("DrugUnit", "TAB"), ("DrugValue", "1234.5")), false);
        Assert.Equal("1,234.50", values["DrugValue"]);
        Assert.Equal("", values["DrugDesc"]);
    }

    [Fact]
    public void Tooth_reference_category_keeps_the_first_character_of_the_choice()
    {
        var values = CodeTableService.Validate(T("tooth-reference"), V(("ToothRefCode", "11"), ("ToothRefCat", "AB"), ("ToothRefDesc", "Incisor")), false);
        Assert.Equal("A", values["ToothRefCat"]);
    }

    [Fact]
    public void Editing_keeps_an_existing_short_code()
    {
        // Existing data has codes the 3-character rule would refuse (e.g. CheckItems "BILL"); Edit never re-checked the code.
        var values = CodeTableService.Validate(T("check-items"), V(("CatCode", "BI"), ("CatName", "Bill"), ("Amount", "30")), editing: true);
        Assert.Equal("BI", values["CatCode"]);
    }

    [Fact]
    public void Identity_key_tables_do_not_ask_for_the_key_on_add()
    {
        var values = CodeTableService.Validate(T("decline-reason"), V(("DeclineReason", "Not covered"), ("Remarks", "")), false);
        Assert.False(values.ContainsKey("ID"));
        Assert.Equal("Not covered", values["DeclineReason"]);
    }

    [Fact]
    public void Catalog_identifiers_are_safe_and_every_table_has_one_key()
    {
        foreach (var t in CodeTableCatalog.All)
        {
            Assert.Single(t.Fields, f => f.IsKey);
            foreach (var name in t.Fields.Select(f => f.Column).Append(t.Table).Append(t.AuditUserColumn).Append(t.AuditDateColumn).OfType<string>()
                         .Concat(t.DeleteChecks.SelectMany(c => new[] { c.Table, c.Column })))
                Assert.Matches(CodeTableRepository.SafeName(), name);
            if (t.ListReportKey is not null) Assert.NotNull(ListReportCatalog.Find(t.ListReportKey));
        }
        Assert.Equal(CodeTableCatalog.All.Count, CodeTableCatalog.All.Select(t => t.Key).Distinct().Count());
    }

    [Fact]
    public void List_report_catalog_has_the_26_desktop_reports_with_unique_keys()
    {
        Assert.Equal(26, ListReportCatalog.All.Count);
        Assert.Equal(26, ListReportCatalog.All.Select(r => r.Key).Distinct().Count());
        Assert.All(ListReportCatalog.All, r => Assert.StartsWith("cis.ListReport_", r.Procedure));
    }
}

public class CodeTableServiceTests
{
    private sealed class FakeRepository : ICodeTableRepository
    {
        public List<Dictionary<string, string?>> Rows { get; } = [];
        public IReadOnlyDictionary<string, string>? LastFilters { get; private set; }

        public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, CodeTableDefinition table, IReadOnlyDictionary<string, string> filters, int max)
        {
            LastFilters = filters;
            return Task.FromResult(Rows.Select(r => r.ToDictionary(kv => kv.Key, kv => (object?)kv.Value)).ToList());
        }

        public Task<bool> ExistsAsync(IDbSession session, CodeTableDefinition table, string code) => Task.FromResult(Rows.Any(r => r[table.KeyField.Column] == code));

        public Task InsertAsync(IDbSession session, CodeTableDefinition table, IReadOnlyDictionary<string, string?> values, string userCode, DateTime date)
        {
            Rows.Add(new Dictionary<string, string?>(values) { ["user"] = userCode, ["date"] = date.ToString("yyyy-MM-dd") });
            return Task.CompletedTask;
        }

        public Task<int> UpdateFirstAsync(IDbSession session, CodeTableDefinition table, string code, IReadOnlyDictionary<string, string?> values, string userCode, DateTime date)
        {
            var row = Rows.FirstOrDefault(r => r[table.KeyField.Column] == code);
            if (row is null) return Task.FromResult(0);
            foreach (var (k, v) in values) row[k] = v;
            return Task.FromResult(1);
        }

        public Task<int> DeleteAsync(IDbSession session, CodeTableDefinition table, string code) =>
            Task.FromResult(Rows.RemoveAll(r => r[table.KeyField.Column] == code));

        public Task<List<LookupOption>> ChoicesAsync(IDbSession session, ChoiceSource source) => Task.FromResult(new List<LookupOption> { new("A", "Anterior") });

        public HashSet<string> UsedIn { get; } = [];

        public Task<bool> UsedAsync(IDbSession session, UsageCheck check, string code) => Task.FromResult(UsedIn.Contains(check.Table));
    }

    private const string AllRights = "1111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111";
    private readonly FakeSessions _sessions = new();
    private readonly FakeRepository _repository = new();
    private readonly CodeTableService _service;

    public CodeTableServiceTests() => _service = new CodeTableService(_sessions, _repository, new FixedClock(new DateTime(2026, 11, 1, 10, 0, 0)));

    private static CodeTableSaveRequest R(params (string, string?)[] values) => new(values.ToDictionary(v => v.Item1, v => v.Item2));

    [Fact]
    public async Task Health_type_is_upper_cased_and_delete_runs_the_usage_checks_in_order()
    {
        Assert.Equal("Record Saved...", (await _service.CreateAsync("health-type", R(("HLTCode", "z"), ("HLTDesc", "zz test")), "DEV", AllRights)).Message);
        Assert.Equal(("Z", "ZZ TEST"), (_repository.Rows[^1]["HLTCode"], _repository.Rows[^1]["HLTDesc"]));
        Assert.Equal("Please Enter Health Code", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.DeleteAsync("health-type", " ", AllRights))).Message);
        _repository.UsedIn.UnionWith(["PLN", "PRM"]);
        Assert.Equal("This record cannot be deleted because it is used in Plan table!",
            (await Assert.ThrowsAsync<ConflictException>(() => _service.DeleteAsync("health-type", "Z", AllRights))).Message);
        _repository.UsedIn.Clear();
        Assert.Equal("Record Deleted... ", (await _service.DeleteAsync("health-type", "Z", AllRights)).Message);
    }

    [Fact]
    public async Task Add_saves_with_the_user_and_date_and_refuses_a_duplicate_code()
    {
        Assert.Equal(CodeTableService.Saved, (await _service.CreateAsync("fee", R(("FeeCode", "F01"), ("FeeDesc", "Fee"), ("FeeRemark", "")), "DEVADMIN", AllRights)).Message);
        Assert.Equal(("DEVADMIN", "2026-11-01"), (_repository.Rows[0]["user"], _repository.Rows[0]["date"]));
        Assert.True(_sessions.Last.Committed);

        var e = await Assert.ThrowsAsync<ConflictException>(() => _service.CreateAsync("fee", R(("FeeCode", "F01"), ("FeeDesc", "Other")), "DEVADMIN", AllRights));
        Assert.Equal(CodeTableService.Duplicated, e.Message);
        Assert.Single(_repository.Rows);
    }

    [Fact]
    public async Task Edit_and_delete_answer_the_desktop_texts()
    {
        await _service.CreateAsync("repudiation", R(("RepCode", "R1"), ("RepDesc", "Old")), "U", AllRights);
        Assert.Equal(CodeTableService.Updated, (await _service.UpdateAsync("repudiation", "R1", R(("RepDesc", "New")), "U", AllRights)).Message);
        Assert.Equal("New", _repository.Rows[0]["RepDesc"]);
        Assert.Equal(CodeTableService.Deleted, (await _service.DeleteAsync("repudiation", "R1", AllRights)).Message);
        Assert.Empty(_repository.Rows);
        await Assert.ThrowsAsync<NotFoundException>(() => _service.DeleteAsync("repudiation", "R1", AllRights));
    }

    [Fact]
    public async Task Search_uses_the_first_three_characters_like_the_desktop_and_ignores_unknown_filters()
    {
        await _service.SearchAsync("procedures", new Dictionary<string, string> { ["ProcedureCode"] = "ABCD", ["Hack"] = "x" }, AllRights);
        Assert.Equal(new Dictionary<string, string> { ["ProcedureCode"] = "ABC" }, _repository.LastFilters);
    }

    [Fact]
    public async Task Each_table_needs_its_screen_right()
    {
        var onlyCheckItem = new string('0', 142) + "1";
        await _service.SearchAsync("check-items", new Dictionary<string, string>(), onlyCheckItem);
        await Assert.ThrowsAsync<ForbiddenException>(() => _service.SearchAsync("fee", new Dictionary<string, string>(), onlyCheckItem));
        Assert.Throws<NotFoundException>(() => _service.Describe("no-such-table", AllRights));
    }
}

public class ListReportServiceTests
{
    private sealed class FakeRows : IListReportRepository
    {
        public Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, string procedure, string? code) =>
            Task.FromResult(new List<Dictionary<string, object?>> { new() { ["CatCode"] = "BILL", ["CatName"] = "BILL ACCUTE", ["Amount"] = "30" } });
    }

    [Fact]
    public async Task Prints_the_report_with_its_desktop_title_and_checks_the_right()
    {
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new ListReportService(new FakeSessions(), new FakeRows(), exporter, new FixedClock(new DateTime(2026, 11, 1)));

        var csv = await service.PrintAsync("check-items", null, ReportFileType.Csv, new string('0', 142) + "1");
        Assert.Equal("RptCheckList.csv", csv.FileName);
        Assert.Contains("Category Code,Category Name,Amount", Encoding.UTF8.GetString(csv.Content));

        await Assert.ThrowsAsync<ForbiddenException>(() => service.PrintAsync("fee", null, ReportFileType.Pdf, new string('0', 142) + "1"));
        await Assert.ThrowsAsync<NotFoundException>(() => service.PrintAsync("nope", null, ReportFileType.Pdf, "1"));
    }
}

/// <summary>Catalog against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class MaintenanceDatabaseTests
{
    [Fact]
    public async Task Every_catalog_table_and_column_exists_and_every_list_report_runs()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().OpenAsync();
        var conn = s.Sql().Connection;
        foreach (var t in CodeTableCatalog.All)
            foreach (var column in t.Fields.Select(f => f.Column).Append(t.AuditUserColumn).Append(t.AuditDateColumn).OfType<string>())
                Assert.True(await conn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM sys.columns WHERE object_id = OBJECT_ID(@t) AND name = @c", new { t = "dbo." + t.Table, c = column }) == 1,
                    $"{t.Table}.{column} does not exist");

        var repository = new ListReportRepository();
        foreach (var report in ListReportCatalog.All)
        {
            var rows = await repository.RowsAsync(s, report.Procedure, null);
            if (rows.Count > 0)
                foreach (var column in report.Columns) Assert.True(rows[0].ContainsKey(column.Key), $"{report.Procedure} has no column {column.Key}");
        }
    }

    [Fact]
    public async Task Usage_checks_find_codes_in_use_and_skip_missing_tables()
    {
        if (TestDb.Cis is null) return;
        await using var s = await TestDb.Sessions().OpenAsync();
        var repo = new CodeTableRepository();
        var used = await s.Sql().Connection.ExecuteScalarAsync<string>("SELECT TOP 1 RTRIM(HLTCode) FROM dbo.PLN WHERE LEN(HLTCode) > 0");
        Assert.True(await repo.UsedAsync(s, new UsageCheck("PLN", "HLTCode", "x"), used!));
        Assert.False(await repo.UsedAsync(s, new UsageCheck("PLN", "HLTCode", "x"), "#"));
        Assert.False(await repo.UsedAsync(s, new UsageCheck("NoSuchTableZZ", "HLTCode", "x"), used!));
    }

    [Fact]
    public async Task Code_table_repository_adds_edits_finds_and_deletes_a_row()
    {
        if (TestDb.Cis is null) return;
        await using var s = await TestDb.Sessions().BeginAsync();
        var table = CodeTableCatalog.Find("check-items")!;
        var repo = new CodeTableRepository();

        await repo.InsertAsync(s, table, new Dictionary<string, string?> { ["CatCode"] = "ZZ9", ["CatName"] = "CIS WEB TEST", ["Amount"] = "1" }, "CISTEST", new DateTime(2026, 11, 1));
        Assert.True(await repo.ExistsAsync(s, table, "ZZ9"));
        Assert.Equal(1, await repo.UpdateFirstAsync(s, table, "ZZ9", new Dictionary<string, string?> { ["CatCode"] = "ZZ9", ["CatName"] = "CIS WEB TEST 2", ["Amount"] = "2" }, "CISTEST", new DateTime(2026, 11, 1)));
        var rows = await repo.SearchAsync(s, table, new Dictionary<string, string> { ["CatCode"] = "ZZ9" }, 10);
        Assert.Equal("CIS WEB TEST 2", Assert.Single(rows)["CatName"]);
        Assert.Equal(1, await repo.DeleteAsync(s, table, "ZZ9"));
        Assert.NotEmpty(await repo.ChoicesAsync(s, CodeTableCatalog.Find("tooth-reference")!.Fields[1].Choices!));
    }
}
