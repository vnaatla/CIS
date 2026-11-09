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
        services.AddScoped<Application.Maintenance.Entity.IEntityClinicRepository, Maintenance.EntityClinicRepository>();
        services.AddScoped<Application.Maintenance.Entity.IEntityGrpComAccessRepository, Maintenance.EntityGrpComAccessRepository>();
        services.AddScoped<Application.Maintenance.Plan.IPlanCodeRepository, Maintenance.PlanCodeRepository>();
        services.AddScoped<Application.Maintenance.Plan.IPlanPremiumRepository, Maintenance.PlanPremiumRepository>();
        services.AddScoped<Application.Maintenance.Plan.IPlanAnnualLimitRepository, Maintenance.PlanAnnualLimitRepository>();
        services.AddScoped<Application.Maintenance.Plan.IPlanVisitLimitRepository, Maintenance.PlanVisitLimitRepository>();
        services.AddScoped<Application.Maintenance.Plan.IPlanMcoRepository, Maintenance.PlanMcoRepository>();
        services.AddScoped<Application.Membership.IMemberSearchRepository, Membership.MemberSearchRepository>();
        services.AddScoped<Application.Membership.IMbmTopupRepository, Membership.MbmTopupRepository>();
        services.AddScoped<Application.Membership.ICaseTransferRepository, Membership.CaseTransferRepository>();
        services.AddScoped<Application.Membership.IMemberRegistrationRepository, Membership.MemberRegistrationRepository>();
        services.AddScoped<Application.Membership.IGeneralExclusionsRepository, Membership.GeneralExclusionsRepository>();
        services.AddScoped<Application.Membership.IMbmAdjustmentRepository, Membership.MbmAdjustmentRepository>();
        services.AddScoped<Application.Membership.IMbmEnquiryRepository, Membership.MbmEnquiryRepository>();
        services.AddScoped<Application.Invoice.IBatchTransferRepository, Invoice.BatchTransferRepository>();
        services.AddScoped<Application.Invoice.IInvoiceUpdateRepository, Invoice.InvoiceUpdateRepository>();
        services.AddScoped<Application.Invoice.IOpBatchRepository, Invoice.OpBatchRepository>();
        services.AddScoped<Application.Invoice.IInvoicePopupRepository, Invoice.InvoicePopupRepository>();
        services.AddScoped<Application.Invoice.IMmaScheduleRepository, Invoice.MmaScheduleRepository>();
        services.AddScoped<Application.Invoice.IWorksheetRepository, Invoice.WorksheetRepository>();
        services.AddScoped<Application.Invoice.IMemberSuppRepository, Invoice.MemberSuppRepository>();
        services.AddScoped<Application.Invoice.IClinicInvoiceRepository, Invoice.ClinicInvoiceRepository>();
        services.AddScoped<Application.Invoice.IMemberInvoiceRepository, Invoice.MemberInvoiceRepository>();
        services.AddScoped<Application.Invoice.IInvoiceEnquiryRepository, Invoice.InvoiceEnquiryRepository>();
        services.AddScoped<Application.Invoice.IInvoiceAdjustmentRepository, Invoice.InvoiceAdjustmentRepository>();
        services.AddScoped<Application.Audit.IAuditReviewRepository, Audit.AuditReviewRepository>();
        services.AddScoped<Application.Audit.IAuditTrackingRepository, Audit.AuditTrackingRepository>();
        services.AddScoped<Application.Lg.ILgRepository, Lg.LgRepository>();
        services.AddScoped<Application.Mco.IMcoListRepository, Mco.McoListRepository>();
        services.AddScoped<Application.Bordx.IBordxWsRepository, Bordx.BordxWsRepository>();
        services.AddScoped<Application.Bordx.IBordxBsnRepository, Bordx.BordxBsnRepository>();
        services.AddScoped<Application.Bordx.IBordxInvRepository, Bordx.BordxInvRepository>();
        services.AddScoped<Application.Payment.IPayClinicRepository, Payment.PayClinicRepository>();
        services.AddScoped<Application.AcExport.IPostInvClnRepository, AcExport.PostInvClnRepository>();
        services.AddScoped<Application.AcExport.IPostInvClnGstRepository, AcExport.PostInvClnGstRepository>();
        services.AddScoped<Application.Payment.IPayClnPostRepository, Payment.PayClnPostRepository>();
        services.AddScoped<Application.Payment.IPayMbmOldRepository, Payment.PayMbmOldRepository>();
        services.AddScoped<Application.Receipt.IRcpMbmRepository, Receipt.RcpMbmRepository>();
        services.AddScoped<Application.Receipt.IRcpMnrbRepository, Receipt.RcpMnrbRepository>();
        services.AddScoped<Application.Receipt.IRcpUpRepository, Receipt.RcpUpRepository>();
        services.AddScoped<Application.Receipt.IRcpFfsRepository, Receipt.RcpFfsRepository>();
        services.AddScoped<Application.Receipt.IRcpPayRepository, Receipt.RcpPayRepository>();
        services.AddScoped<Application.Payment.IPayFfsRepository, Payment.PayFfsRepository>();
        services.AddScoped<Application.Payment.IPayTrackingRepository, Payment.PayTrackingRepository>();
        services.AddScoped<Application.Payment.IPayMbmRepository, Payment.PayMbmRepository>();
        services.AddScoped<Application.Payment.IPayClnInvRepository, Payment.PayClnInvRepository>();
        services.AddScoped<Application.Bordx.IPrintBordxRepository, Bordx.PrintBordxRepository>();
        services.AddScoped<Application.Bordx.IBordxMonitoringRepository, Bordx.BordxMonitoringRepository>();
        services.AddScoped<Application.Bordx.IBordxDeliveryRepository, Bordx.BordxDeliveryRepository>();
        services.AddScoped<Application.Bordx.IBordxDateRepository, Bordx.BordxDateRepository>();
        services.AddScoped<Application.Bordx.IBordxRemoveRepository, Bordx.BordxRemoveRepository>();
        services.AddScoped<Application.Mco.IMcoZeroRepository, Mco.McoZeroRepository>();
        services.AddScoped<Application.Mco.IMcoGstRepository, Mco.McoGstRepository>();
        services.AddScoped<Application.Mco.IMcoInvoiceRepository, Mco.McoInvoiceRepository>();
        services.AddScoped<Application.Mco.IMcoRefundRepository, Mco.McoRefundRepository>();
        services.AddScoped<Application.Lg.ILgSaveRepository, Lg.LgSaveRepository>();
        services.AddSingleton<Application.Documents.ILetterWriter, Documents.PdfLetterWriter>();
        services.AddScoped<Application.Claim.IPassToClaimRepository, Claim.PassToClaimRepository>();
        services.AddScoped<Application.Claim.IUpdateClaimsDataRepository, Claim.UpdateClaimsDataRepository>();
        services.AddSingleton<Application.Documents.ISpreadsheetReader, Documents.SpreadsheetReader>();

        // PDFsharp needs a font source on hosts where it cannot find fonts by itself.
        GlobalFontSettings.FontResolver ??= new PdfFontResolver();
        return services;
    }
}
