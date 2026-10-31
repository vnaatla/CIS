using Cis.Api.Audit;
using Cis.Api.Auth;
using Cis.Api.Auth.Login;
using Cis.Api.Data;
using Cis.Api.Documents;
using Cis.Api.Email;
using Cis.Api.Files;
using Cis.Api.Lookups;

namespace Cis.Api;

/// <summary>Composition root of the CIS services. Each migration phase adds one <c>Add…Screens</c> method here.</summary>
public static class CisServiceCollectionExtensions
{
    /// <summary>Cross-cutting services: database, auth / session, audit, documents, files, e-mail, lookups.</summary>
    public static IServiceCollection AddCisPlatform(this IServiceCollection services, IConfiguration configuration)
    {
        services.Configure<AuthOptions>(configuration.GetSection(AuthOptions.SectionName));
        services.Configure<SessionTrackingOptions>(configuration.GetSection(SessionTrackingOptions.SectionName));
        services.Configure<ReportOptions>(configuration.GetSection(ReportOptions.SectionName));
        services.Configure<Files.FileOptions>(configuration.GetSection(Files.FileOptions.SectionName));
        services.Configure<EmailOptions>(configuration.GetSection(EmailOptions.SectionName));

        services.AddSingleton(TimeProvider.System);
        services.AddMemoryCache();
        services.AddSingleton<DbConnections>();
        services.AddScoped<IDbSessions, DbSessions>();
        services.AddSingleton<TokenService>();

        services.AddSingleton<ILoginAttempts, LoginAttempts>();
        services.AddSingleton<ILoginLog, LoginLogFile>();
        services.AddScoped<ILoginRepository, LoginRepository>();
        services.AddScoped<ILoginService, LoginService>();

        services.AddSingleton<IErrorLog, ErrorLog>();
        services.AddScoped<IActionLogRepository, ActionLogRepository>();
        services.AddScoped<IActionLogService, ActionLogService>();

        services.AddSingleton<IReportTemplateStore, ReportTemplateStore>();
        services.AddSingleton<ExcelReportWriter>();
        services.AddSingleton<PdfReportWriter>();
        services.AddSingleton<IReportExporter, ReportExporter>();
        services.AddSingleton<IFileDrop, FileDrop>();
        services.AddSingleton<IEmailSender, SmtpEmailSender>();

        services.AddScoped<ILookupRepository, LookupRepository>();
        services.AddScoped<ILookupService, LookupService>();
        return services;
    }
}
