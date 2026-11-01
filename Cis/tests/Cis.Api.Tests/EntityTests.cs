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
}
