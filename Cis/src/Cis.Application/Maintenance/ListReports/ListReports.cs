using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Domain.Errors;
using Cis.Domain.Security;

namespace Cis.Application.Maintenance.ListReports;

/// <summary>One of the desktop's maintenance list print-outs (a DataReport).</summary>
/// <param name="Key">URL key, e.g. <c>check-items</c>.</param>
/// <param name="LegacyReport">VB6 DataReport name, e.g. <c>RptCheckList</c>.</param>
/// <param name="Title">Report heading (the DataReport's title label).</param>
/// <param name="Procedure">Stored procedure returning the rows (<c>cis.ListReport_*</c>).</param>
/// <param name="Access">USRAccess position of the screen that prints it.</param>
/// <param name="Columns">Columns in the DataReport's order.</param>
public record ListReportDefinition(string Key, string LegacyReport, string Title, string Procedure, int Access, IReadOnlyList<ReportColumn> Columns);

/// <summary>The 26 DataReports of medixcis.vbp: titles and column captions as in the .Dsr files, data from DataEnvironment1's commands.</summary>
public static class ListReportCatalog
{
    private static ReportColumn T(string key, string header) => new(key, header);
    private static ReportColumn M(string key, string header) => new(key, header, ReportFormat.Money);
    private static ReportColumn D(string key, string header) => new(key, header, ReportFormat.Date);

    /// <summary>All reports.</summary>
    public static readonly IReadOnlyList<ListReportDefinition> All =
    [
        new("check-items", "RptCheckList", "CHECK ITEMS LISTING", "cis.ListReport_CheckItems", AccessRight.MaintenanceCheckItem,
            [T("CatCode", "Category Code"), T("CatName", "Category Name"), T("Amount", "Amount")]),
        new("co-payment", "RptCoPaymentList", "CO-PAYMENT LISTING", "cis.ListReport_CoPayment", AccessRight.MaintenanceCalculation,
            [T("COPCode", "Co-Payment Code"), T("COPDesc", "Description"), T("COPRemark", "Remark")]),
        new("fee", "RptFeeList", "FEE LISTING", "cis.ListReport_Fee", AccessRight.MaintenanceCalculation,
            [T("FeeCode", "Fee Code"), T("FeeDesc", "Description"), T("FeeRemark", "Remark")]),
        new("reason", "RptReasonList", "REASON LISTING", "cis.ListReport_Reason", AccessRight.MaintenanceCase,
            [T("ReasonCode", "Treatment Reason Code"), T("ReasonDesc", "Description")]),
        new("procedures", "rptProcedureslist", "PROCEDURES LISTING", "cis.ListReport_Procedures", AccessRight.MaintenanceCase,
            [T("ProcedureCode", "Procedure Code"), T("ProcedureDesc", "Description")]),
        new("drug", "rptDruglist", "DRUG LISTING", "cis.ListReport_Drug", AccessRight.MaintenanceCase,
            [T("DrugCode", "Drug Code"), T("DrugUnit", "Drug Unit"), T("DrugValue", "Drug Value"), T("DrugDesc", "Drug Description")]),
        new("repudiation", "rptRepudiationList", "REPUDIATION LISTING", "cis.ListReport_Repudiation", AccessRight.MaintenanceCase,
            [T("RepCode", "Repudiation Code"), T("RepDesc", "Description")]),
        new("diagnosis-not-payable", "rptDiagnosisList", "DIAGNOSIS NOT PAYABLE LISTING", "cis.ListReport_DiagnosisNotPayable", AccessRight.MaintenanceCase,
            [T("PLNCode", "Plan Code"), T("DiagCode", "Diagnosis Code")]),
        new("health-type", "RptHltList", "HEALTH LISTING", "cis.ListReport_HealthType", AccessRight.MaintenancePlan,
            [T("HLTCode", "Health Code"), T("HLTDesc", "Description")]),
        new("insured-type", "RptInsList", "INSURED LISTING", "cis.ListReport_InsuredType", AccessRight.MaintenancePlan,
            [T("INSCode", "Insured Code"), T("INSDesc", "Description")]),
        new("mco", "RptMCOList", "MCO LISTING", "cis.ListReport_Mco", AccessRight.MaintenancePlan,
            [T("HLTCode", "Health Code"), T("PLNCode", "Plan Code"), T("INSCode", "Insured Code"), M("MCOAmtAdult", "MCO Fees (Adult)"),
             M("MCOAmtChild", "MCO Fees (Child)"), D("MCOEffDate", "Effective Date")]),
        new("premium", "RptPremiumList", "PREMIUM LISTING", "cis.ListReport_Premium", AccessRight.MaintenancePlan,
            [T("PLNCode", "Plan Code"), M("PRMAmount", "Premium Amount"), T("HLTCode", "Health  Type"), T("INSCode", "Insured  Code"), T("PAYCode", "Payor"),
             M("SuppPRMAmt", "Supplementary Premium Amount"), T("AGECode", "Age Code"), D("PRMEffDate", "Premium Effective Date"),
             T("SuppPRMStatus", "Supplementary Premium Status"), T("PRMVersion", "Premium Version")]),
        new("visit-limit", "RptVisitList", "VISIT LIMIT LISTING", "cis.ListReport_VisitLimit", AccessRight.MaintenancePlan,
            [T("PLNCode", "Plan Code"), T("VisitLimit", "Visit Limit"), T("HLTCode", "Health  Type"), T("INSCode", "Insured  Type"), T("PAYCode", "Payor"),
             T("SuppVisitLimit", "Supplementary Visit Limit"), T("GRPCompany", "Group Company"), D("VisitEffDate", "Visit Limit Effective Date"),
             T("SuppVisitLimitStatus", "Supplementary Visit Limit Status"), T("VisitLimitVersion", "Visit Version")]),
        new("annual-limit", "RptAnnualList", "ANNUAL LIMIT LISTING", "cis.ListReport_AnnualLimit", AccessRight.MaintenancePlan,
            [T("PLNCode", "Plan Code"), M("AnnLimit", "Annual Limit"), T("HLTCode", "Health  Type"), T("INSCode", "Insured  Type"), T("PAYCode", "Payor"),
             M("SuppLimit", "Supplementary Annual Limit"), T("GRPCompany", "Group Company"), D("AnnEffDate", "Annual Limit Effective Date"),
             T("SuppLimitStatus", "Suppl. Annual Limit Status"), T("AnnLimitVersion", "Annual Version")]),
        new("co-payment-member", "RptCOPayMBM", "CO-PAYMENT MBM LISTING", "cis.ListReport_CoPaymentMember", AccessRight.MaintenancePlan,
            [M("CoAmount", "Co-Payment Amount"), T("HLTType", "Health Code"), T("PLNCode", "Plan Code")]),
        new("plan", "RptPlanList", "PLAN LISTING", "cis.ListReport_Plan", AccessRight.MaintenancePlan,
            [T("PLNCode", "Plan Code"), T("PLNDesc", "Plan Description"), T("HLTCode", "Health Type"), T("PAYCode", "CIS Payor"),
             T("GRPCompany", "Group Company"), T("HISPAYCode", "HIS Payor")]),
        new("group-clinic", "RptGrpClinicList", "GROUP CLINIC LISTING", "cis.ListReport_GroupClinic", AccessRight.MaintenanceEntity,
            [T("GRPClnCode", "Clinic Group Code"), T("GRPClnName", "Clinic Group Name")]),
        new("doctor", "RptDoctorList", "DOCTOR LISTING", "cis.ListReport_Doctor", AccessRight.MaintenanceEntity,
            [T("DocCode", "Doctor Code"), T("DocName", "Doctor Name")]),
        new("group-company", "RptGrpCompanyList", "GROUP COMPANY LISTING", "cis.ListReport_GroupCompany", AccessRight.MaintenanceEntity,
            [T("GRPCoCode", "Group Company Code"), T("GRPCoName", "Group Company Name"), T("GRPCoPaycode", "Payor Code")]),
        new("group-company-clinic", "RptGrpComCLN", "GROUP COMPANY CLINIC LISTING", "cis.ListReport_GroupCompanyClinic", AccessRight.MaintenanceEntity,
            [T("GrpComCode", "Grp Com Code"), T("GRPCoName", "Group Company Name"), T("GrpComClnCode", "Clinic Code"), T("CLNName", "Clinic Name")]),
        new("tooth-category", "RptToothCat", "TOOTH CATERGORY LISTING", "cis.ListReport_ToothCategory", AccessRight.MaintenanceBenefitMember,
            [T("ToothRefCatCode", "Catergory Code"), T("ToothRefCatDesc", "Description")]),
        new("tooth-reference", "RptToothRef", "TOOTH REFERENCE LISTING", "cis.ListReport_ToothReference", AccessRight.MaintenanceBenefitMember,
            [T("ToothRefCode", "Catergory Code"), T("ToothRefDesc", "Description"), T("ToothRefCatDesc", "Catergory")]),
        new("dental-benefit-category", "RptDentalBNFCat", "DENTAL BENEFIT CATEROGY LISTING", "cis.ListReport_DentalBenefitCategory", AccessRight.MaintenanceBenefitMember,
            [T("DentalBNFCATCode", "Catergory Code"), T("DentalBNFCATDesc", "Description")]),
        new("dental-benefit-entry", "RptDentalBNFEntry", "DENTAL BENEFIT ENTRY LISTING", "cis.ListReport_DentalBenefitEntry", AccessRight.MaintenanceBenefitMember,
            [T("BNFCode", "BNF Code"), T("DentalBNFCATDesc", "Description"), T("BNFCalculation", "Calculation"), T("BNFNoOfDays", "No of Visit Per Day"),
             T("BNFNoOfAnn", "No of Visit Per Annum"), M("BNFRMPerVisit", "RM Per Visit"), T("BNFPercentage", "Percentage")]),
        new("benefit-category", "RptBNFCat", "BNF CATEGORY LISTING", "cis.ListReport_BenefitCategory", AccessRight.MaintenanceBenefitMember,
            [T("BNFCATCode", "BNF Category Code"), T("BNFCATDescription", "Description")]),
        new("clinic-benefit", "RptClinicBNF", "CLINIC BENEFIT LISTING", "cis.ListReport_ClinicBenefit", AccessRight.MaintenanceBenefitClinic,
            [T("BNFCode", "BNF Code"), T("BNFCATDescription", "Description"), T("CLNBNFType", "Card Type"), M("CLNBNFRMPerVisit", "RM Per Visit"),
             D("CLNBNFEffDate", "Effective Date")]),
    ];

    /// <summary>The report with <paramref name="key"/>, or null.</summary>
    public static ListReportDefinition? Find(string key) => All.FirstOrDefault(r => string.Equals(r.Key, key, StringComparison.OrdinalIgnoreCase));
}

/// <summary>Data access of the list reports.</summary>
public interface IListReportRepository
{
    /// <summary>Rows of <paramref name="procedure"/>, optionally filtered on the report's code.</summary>
    Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, string procedure, string? code);
}

/// <summary>Prints the maintenance list reports (the screens' Print buttons).</summary>
public interface IListReportService
{
    /// <summary>The report as PDF (or another <paramref name="type"/>) for a user with the screen's right.</summary>
    Task<FileOutput> PrintAsync(string key, string? code, ReportFileType type, string? userAccess);
}

/// <inheritdoc />
public sealed class ListReportService(IDbSessions sessions, IListReportRepository repository, IReportExporter exporter, TimeProvider clock) : IListReportService
{
    /// <inheritdoc />
    public async Task<FileOutput> PrintAsync(string key, string? code, ReportFileType type, string? userAccess)
    {
        var report = ListReportCatalog.Find(key) ?? throw new NotFoundException($"Unknown list report '{key}'.");
        if (!AccessRight.Has(userAccess, report.Access)) throw new ForbiddenException();

        await using var session = await sessions.OpenAsync();
        var rows = await repository.RowsAsync(session, report.Procedure, string.IsNullOrWhiteSpace(code) ? null : code.Trim());
        var document = new ReportDocument(report.Title, report.Columns, rows)
        {
            FileName = report.LegacyReport,
            Subtitles = [$"Printed {clock.GetLocalNow():dd/MM/yyyy HH:mm}", $"Total records: {rows.Count}"],
            Landscape = report.Columns.Count > 4,
        };
        return exporter.Export(document, type);
    }
}
