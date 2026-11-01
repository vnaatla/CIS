using Cis.Application.Audit;
using Cis.Application.Auth.Login;
using Cis.Application.Lookups;
using FluentValidation;
using Microsoft.Extensions.DependencyInjection;

namespace Cis.Application;

/// <summary>Registers the use cases (services) and validators. Each migration phase adds its screens here.</summary>
public static class DependencyInjection
{
    /// <summary>Services and validators of the application layer.</summary>
    public static IServiceCollection AddCisApplication(this IServiceCollection services)
    {
        services.AddSingleton(TimeProvider.System);
        services.AddValidatorsFromAssemblyContaining(typeof(DependencyInjection));

        services.AddScoped<ILoginService, LoginService>();
        services.AddScoped<IActionLogService, ActionLogService>();
        services.AddScoped<ILookupService, LookupService>();

        // Phase 1: Maintenance
        services.AddScoped<Maintenance.ListReports.IListReportService, Maintenance.ListReports.ListReportService>();
        services.AddScoped<Maintenance.CodeTables.ICodeTableService, Maintenance.CodeTables.CodeTableService>();
        services.AddScoped<Maintenance.LgTemplates.ILgBindingService, Maintenance.LgTemplates.LgBindingService>();
        services.AddScoped<Maintenance.BenefitClinic.IBenefitClinicService, Maintenance.BenefitClinic.BenefitClinicService>();
        services.AddScoped<Maintenance.BenefitMember.IBenefitMemberService, Maintenance.BenefitMember.BenefitMemberService>();
        services.AddScoped<Maintenance.Entity.IEntityGroupDoctorService, Maintenance.Entity.EntityGroupDoctorService>();
        services.AddScoped<Maintenance.Entity.IEntityClinicService, Maintenance.Entity.EntityClinicService>();
        return services;
    }
}
