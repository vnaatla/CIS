using Cis.Application.Abstractions;
using Cis.Application.Audit;
using Cis.Application.Auth.Login;
using Cis.Application.Documents;
using Cis.Application.Email;
using Cis.Application.Files;
using Cis.Application.Lookups;
using Cis.Infrastructure.Audit;
using Cis.Infrastructure.Auth;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Email;
using Cis.Infrastructure.Files;
using Cis.Infrastructure.Lookups;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PdfSharp.Fonts;

namespace Cis.Infrastructure;

/// <summary>Registers the adapters that implement the application's ports. Each migration phase adds its repositories here.</summary>
public static class DependencyInjection
{
    /// <summary>Database, repositories, documents, files, e-mail, login tracking, logs; options bound from <paramref name="configuration"/>.</summary>
    public static IServiceCollection AddCisInfrastructure(this IServiceCollection services, IConfiguration configuration)
    {
        services.Configure<SessionTrackingOptions>(configuration.GetSection(SessionTrackingOptions.SectionName));
        services.Configure<ReportOptions>(configuration.GetSection(ReportOptions.SectionName));
        services.Configure<FileDropOptions>(configuration.GetSection(FileDropOptions.SectionName));
        services.Configure<EmailOptions>(configuration.GetSection(EmailOptions.SectionName));

        services.AddMemoryCache();
        services.AddSingleton<DbConnections>();
        services.AddScoped<IDbSessions, SqlDbSessions>();

        services.AddSingleton<ILoginAttempts, LoginAttempts>();
        services.AddSingleton<ILoginLog, LoginLogFile>();
        services.AddScoped<ILoginRepository, LoginRepository>();

        services.AddSingleton<IErrorLog, ErrorLog>();
        services.AddScoped<IActionLogRepository, ActionLogRepository>();

        services.AddSingleton<IReportTemplateStore, ReportTemplateStore>();
        services.AddSingleton<ExcelReportWriter>();
        services.AddSingleton<PdfReportWriter>();
        services.AddSingleton<IReportExporter, ReportExporter>();
        services.AddSingleton<IFileDrop, FileDrop>();
        services.AddSingleton<IEmailSender, SmtpEmailSender>();

        services.AddScoped<ILookupRepository, LookupRepository>();

        // Phase 1: Maintenance
        services.AddScoped<Application.Maintenance.ListReports.IListReportRepository, Maintenance.ListReportRepository>();
        services.AddScoped<Application.Maintenance.CodeTables.ICodeTableRepository, Maintenance.CodeTableRepository>();
        services.AddScoped<Application.Maintenance.LgTemplates.ILgBindingRepository, Maintenance.LgBindingRepository>();
        services.AddScoped<Application.Maintenance.BenefitClinic.IBenefitClinicRepository, Maintenance.BenefitClinicRepository>();
        services.AddScoped<Application.Maintenance.BenefitMember.IBenefitMemberRepository, Maintenance.BenefitMemberRepository>();
        services.AddScoped<Application.Maintenance.Entity.IEntityGroupDoctorRepository, Maintenance.EntityGroupDoctorRepository>();

        // PDFsharp needs a font source on hosts where it cannot find fonts by itself.
        GlobalFontSettings.FontResolver ??= new PdfFontResolver();
        return services;
    }
}
