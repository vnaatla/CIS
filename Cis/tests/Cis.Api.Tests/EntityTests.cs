using Cis.Application.Maintenance.Entity;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Maintenance;

namespace Cis.Api.Tests;

public class EntityGroupDoctorServiceTests
{
    private static readonly ClinicGroupDto Valid = new(null, "Klinik Test", "kt", "Jalan 1", "", null, " ali ", null, null, "0312", null, null, null,
        "01/02/2020", "Y - Yes", "", "N - No");

    private readonly EntityGroupDoctorService _service = new(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));

    [Theory]
    [InlineData("", "N", "A", "01/01/2020", EntityGroupDoctorService.GroupName)]
    [InlineData("I", "", "A", "01/01/2020", EntityGroupDoctorService.GroupName)]
    [InlineData("I", "N", "", "01/01/2020", EntityGroupDoctorService.GroupAddress)]
    [InlineData("I", "N", "A", "__/__/____", EntityGroupDoctorService.GroupEffDate)]
    public async Task Group_save_checks_in_the_desktop_order(string initial, string name, string add1, string eff, string message)
    {
        var g = Valid with { Initial = initial, Name = name, Add1 = add1, EffDate = eff };
        Assert.Equal(message, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveGroupAsync(new(g), false, "DEV"))).Message);
    }

    [Fact]
    public async Task Group_edit_does_not_need_the_locked_initial()
    {
        var ex = await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveGroupAsync(new(Valid with { Initial = null, Add1 = "" }), true, "DEV"));
        Assert.Equal(EntityGroupDoctorService.GroupAddress, ex.Message);
    }

    [Fact]
    public void Group_values_are_written_like_SaveClinicGrp()
    {
        var c = EntityGroupDoctorService.CleanGroup(Valid with { Code = "K0001" });
        Assert.Equal(("KLINIK TEST", "KT", "JALAN 1", null, "ALI", "0312", ""), (c.Name, c.Initial, c.Add1, c.Add2, c.Contact, c.Phone, c.Homepage));
        Assert.Equal(("2020-02-01", "Y", "", "N"), (c.EffDate, c.Tax, c.Conversion, c.PrePost));
        Assert.Null(EntityGroupDoctorService.CleanGroup(Valid with { EffDate = "31/02/2020" }).EffDate);
    }

    [Fact]
    public async Task Doctor_needs_a_name_and_delete_a_code()
    {
        Assert.Equal(EntityGroupDoctorService.DoctorName, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveDoctorAsync(new(null, ""), false, "DEV"))).Message);
        Assert.Equal(EntityGroupDoctorService.DoctorCode, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.DeleteDoctorAsync(" "))).Message);
        Assert.Equal(EntityGroupDoctorService.GroupCode, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.DeleteGroupAsync(null))).Message);
    }
}

public class EntityClinicServiceTests
{
    internal static readonly ClinicDto Valid = new(null, "Klinik Web", "zk", "A0001 - AMAR", "G", "1", "Jalan ZZ 1", "", null, "50000",
        "Kuala Lumpur", "W.P. Kuala Lumpur", "01/02/2020", "Y", "N", null, "2", "0312", null, null, null, null, null, null, null, null, null,
        "Y", null, "N", "", "Jalan", "9-5", null, null, null, null, "", null, "R1");

    private readonly EntityClinicService _service = new(new FakeSessions(), null!, null!, new FixedClock(new DateTime(2026, 11, 1)));

    [Fact]
    public async Task Save_checks_run_in_the_desktop_order()
    {
        var empty = new ClinicDto(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null,
            null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
        Assert.Equal("Please Enter Clinic Name", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(empty, false, "DEV"))).Message);
        Assert.Equal(" Please Enter FFS Percentage", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { Ffs = "Y", FfsPerc = "0" }, false, "DEV"))).Message);
        Assert.Equal(" Please Operating hours", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { OperatingHr = "" }, false, "DEV"))).Message);
        Assert.Equal("Please Enter Received Date!", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { ReceivedDate = "__/__/____" }, false, "DEV"))).Message);
    }

    [Fact]
    public void Values_and_filters_follow_the_desktop()
    {
        var c = EntityClinicService.Clean(Valid with { Code = "Z123" });
        Assert.Equal(("KLINIK WEB", "A0001", "G", "1", "", "KUALA LUMPUR", "2020-02-01", "2", null, "", "R1"),
            (c.Name, c.GroupCode, c.Type, c.BnfType, c.Add2, c.City, c.ReceivedDate, c.NoOfDoc, c.FfsPerc, c.Online, c.RegNo));
        var f = EntityClinicService.ReduceFilters(Valid);
        Assert.Equal(("ZK", "KLINIK WEB", "A0001", "Y", null), (f.Initial, f.Name, f.GroupCode, f.Panel, f.Add1));
    }

    [Fact]
    public async Task Delete_checks_name_then_selection() =>
        Assert.Equal("Please Enter Clinic Name", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.DeleteAsync(null, ""))).Message);
}

public class EntityGrpComAccessServiceTests
{
    private readonly EntityGrpComAccessService _service = new(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));

    [Theory]
    [InlineData("B006 - BANK SIMPANAN", "B006")]
    [InlineData("A001X1234567", "A001X12345")]
    [InlineData(" ", null)]
    public void Combo_codes_are_cut_like_the_desktop(string text, string? code) => Assert.Equal(code, EntityGrpComAccessService.CodeOf(text));

    [Theory]
    [InlineData("", "01/01/2026", "31/12/2026", EntityGrpComAccessService.ClinicFirst)]
    [InlineData("A001", "__/__/____", "31/12/2026", EntityGrpComAccessService.FromInvalid)]
    [InlineData("A001", "01/01/2026", "", EntityGrpComAccessService.ToInvalid)]
    [InlineData("A001", "01/02/2026", "01/01/2026", EntityGrpComAccessService.FromAfterTo)]
    [InlineData("A001", "01/01/2026", "01/01/2099", EntityGrpComAccessService.ToInvalid)]
    public async Task Access_save_checks_in_the_desktop_order(string code, string from, string to, string message) =>
        Assert.Equal(message, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAccessAsync(new(code, from, to), "DEV"))).Message);

    [Fact]
    public async Task Group_company_clinic_and_clinic_search_need_their_boxes()
    {
        Assert.Equal(EntityGrpComAccessService.GroupCompanyFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(new("", "A001", null, null), false, "DEV"))).Message);
        Assert.Equal(EntityGrpComAccessService.ClinicFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(new("B006", "", null, null), false, "DEV"))).Message);
        Assert.Equal(EntityGrpComAccessService.NameToSearch, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.FindClinicsAsync(""))).Message);
    }
}

/// <summary>Clinic Group / Doctor procedures against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class EntityDatabaseTests
{
    [Fact]
    public async Task Group_code_is_generated_free_saved_found_and_deleted()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new EntityGroupDoctorRepository();
        var code = await repo.NextGroupCodeAsync(s, "B");
        Assert.Matches("^B[0-9]{4}$", code);
        Assert.Null(await repo.GetGroupAsync(s, code));
        var g = EntityGroupDoctorService.CleanGroup(new(code, "ZZ CIS WEB TEST", "BZZ", "ADDR", null, null, null, null, null, null, null, null, null,
            "01/01/2020", "Y", "N", "N"));
        Assert.Equal(1, await repo.SaveGroupAsync(s, true, g, "CISTEST", new DateTime(2026, 11, 1)));
        Assert.NotEqual(code, await repo.NextGroupCodeAsync(s, "B"));
        Assert.Equal(("ZZ CIS WEB TEST", "01/01/2020", "Y"), ((await repo.GetGroupAsync(s, code))!.Name, (await repo.GetGroupAsync(s, code))!.EffDate,
            (await repo.GetGroupAsync(s, code))!.Tax));
        Assert.Contains(await repo.SearchGroupsAsync(s, null, "ZZ CIS WEB"), r => r.Code == code);
        Assert.False(await repo.GroupInUseAsync(s, code));
        Assert.Equal(1, await repo.DeleteGroupAsync(s, code));
    }

    [Fact]
    public async Task Doctor_code_is_generated_saved_found_and_deleted()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new EntityGroupDoctorRepository();
        var code = await repo.NextDoctorCodeAsync(s, "Z");
        Assert.Matches("^Z[0-9]{5}$", code);
        Assert.Equal(1, await repo.SaveDoctorAsync(s, true, code, "ZZ CIS WEB DOCTOR", "CISTEST", new DateTime(2026, 11, 1)));
        Assert.True(await repo.DoctorNameExistsAsync(s, "ZZ CIS WEB DOCTOR"));
        Assert.Contains(await repo.SearchDoctorsAsync(s, code, null), r => r.Name == "ZZ CIS WEB DOCTOR");
        Assert.Equal(1, await repo.DeleteDoctorAsync(s, code));
    }

    [Fact]
    public async Task Clinic_code_is_generated_saved_found_and_deleted()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new EntityClinicRepository();
        var setup = await repo.SetupAsync(s);
        Assert.NotEmpty(setup.Groups);
        Assert.NotEmpty(setup.States);
        var code = await repo.NextCodeAsync(s, "Z");
        Assert.StartsWith("Z", code);
        Assert.Empty(await repo.SearchAsync(s, EntityClinicService.ReduceFilters(EntityClinicServiceTests.Valid with { Code = code, Initial = null, Name = null,
            State = null, City = null, PostCode = null, Panel = null, GroupCode = null, Type = null, Hrs24 = null, Ffs = null })));
        Assert.False(await repo.AddressExistsAsync(s, "JALAN ZZ CIS WEB 1"));
        var row = EntityClinicService.Clean(EntityClinicServiceTests.Valid with { Code = code, Add1 = "Jalan ZZ CIS WEB 1", Ffs = "Y", FfsPerc = "12.6" });
        Assert.Equal(1, await repo.SaveAsync(s, true, row, "CISTEST", new DateTime(2026, 11, 1)));
        Assert.True(await repo.AddressExistsAsync(s, "JALAN ZZ CIS WEB 1"));
        var found = Assert.Single(await repo.SearchAsync(s, EntityClinicService.ReduceFilters(EntityClinicServiceTests.Valid with { Code = code, Ffs = "Y" })));
        Assert.Equal(("KLINIK WEB", 13m), ((string)found["CLNName"]!, (decimal)found["CLNMCOProcfeePerc"]!));
        Assert.False(await repo.HasCasesAsync(s, code));
        Assert.Equal(1, await repo.DeleteAsync(s, code));
    }

    [Fact]
    public async Task Group_company_clinic_and_clinic_access_rows_are_saved_found_and_removed()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new EntityGrpComAccessRepository();
        var lookups = await repo.LookupsAsync(s);
        Assert.NotEmpty(lookups.GroupCompanies);
        var clinic = lookups.Clinics[0].Code;
        Assert.False(await repo.ExistsAsync(s, "ZZGRP", clinic));
        Assert.Equal(1, await repo.SaveAsync(s, true, null, null, "ZZGRP", clinic, "CISTEST", new DateTime(2026, 11, 1)));
        Assert.Equal(1, await repo.SaveAsync(s, false, "ZZGRP", clinic, "ZZGRQ", clinic, "CISTEST", new DateTime(2026, 11, 1)));
        Assert.True(await repo.ExistsAsync(s, "ZZGRQ", clinic));
        Assert.Equal(1, await repo.DeleteAsync(s, "ZZGRQ", clinic));
        var page = await repo.SearchAsync(s, null, clinic, 5);
        Assert.True(page.Rows.Count <= 5 && page.Total >= page.Rows.Count);

        Assert.NotNull(await repo.ClinicNameAsync(s, clinic));
        Assert.Null(await repo.ClinicNameAsync(s, "ZZ#99"));
        var now = new DateTime(2078, 1, 1, 10, 0, 0);
        await repo.InsertAccessAsync(s, clinic, new DateTime(2078, 1, 1), new DateTime(2078, 12, 31), "CISTEST", now);
        Assert.Equal("CISTEST", (await repo.AccessListAsync(s))[0].CreatedBy);
        Assert.NotEmpty(await repo.FindClinicsAsync(s, "KLINIK"));
    }
}
