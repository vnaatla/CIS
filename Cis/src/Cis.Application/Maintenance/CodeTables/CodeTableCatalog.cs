using Cis.Domain.Security;

namespace Cis.Application.Maintenance.CodeTables;

/// <summary>How a field is entered and stored.</summary>
public enum FieldKind
{
    /// <summary>Free text.</summary>
    Text,
    /// <summary>Digits and one decimal point (Statndard.KeyNumber), stored as typed.</summary>
    Number,
    /// <summary>Number stored as text formatted <c>#,##0.00</c> (frmMaintenanceOthers Drug Value: <c>Format(txt, "#,##0.00")</c>).</summary>
    MoneyText,
    /// <summary>Picked from <see cref="CodeField.Choices"/>; the desktop stored the first character of the combo text.</summary>
    ChoiceFirstChar,
}

/// <summary>How the desktop's Search used a field (the form's strAppend).</summary>
public enum SearchMatch
{
    /// <summary>Not searchable.</summary>
    None,
    /// <summary><c>Column = 'value'</c>.</summary>
    EqualsValue,
    /// <summary><c>Column = Trim(Mid(value, 1, 3))</c>.</summary>
    EqualsFirst3,
    /// <summary><c>Column = Mid(value, 1, 1)</c>.</summary>
    EqualsFirstChar,
    /// <summary><c>Column Like '%value%'</c>.</summary>
    Contains,
}

/// <summary>A lookup table a choice field reads its options from.</summary>
/// <param name="Table">Table in CISDB.</param>
/// <param name="CodeColumn">Code column.</param>
/// <param name="NameColumn">Description column.</param>
public record ChoiceSource(string Table, string CodeColumn, string NameColumn);

/// <summary>One field of a code table.</summary>
/// <param name="Column">Column in the table.</param>
/// <param name="Label">Desktop caption.</param>
/// <param name="MaxLength">Text box MaxLength.</param>
/// <param name="Required">Message when empty (null = optional).</param>
/// <param name="MinLength">Minimum length (0 = none).</param>
/// <param name="MinMessage">Message when shorter than <paramref name="MinLength"/>.</param>
/// <param name="Kind">Entry / storage kind.</param>
/// <param name="Search">How Search uses the field.</param>
/// <param name="IsKey">The record's code (duplicate check, edit / delete key).</param>
public record CodeField(string Column, string Label, int MaxLength, string? Required = null, int MinLength = 0, string? MinMessage = null,
    FieldKind Kind = FieldKind.Text, SearchMatch Search = SearchMatch.None, bool IsKey = false)
{
    /// <summary>Options source for <see cref="FieldKind.ChoiceFirstChar"/>.</summary>
    public ChoiceSource? Choices { get; init; }
}

/// <summary>A desktop maintenance tab that edits one simple table (search, list, add, edit, delete, print).</summary>
/// <param name="Key">URL key.</param>
/// <param name="LegacyForm">VB6 form.</param>
/// <param name="Tab">Tab caption (or the form caption when it has no tabs).</param>
/// <param name="Access">USRAccess position of the screen.</param>
/// <param name="Table">Table in CISDB.</param>
/// <param name="AuditUserColumn">Column that receives the user code (<c>Logon</c>).</param>
/// <param name="AuditDateColumn">Column that receives the date (<c>Date</c>).</param>
/// <param name="Fields">Fields in the form's order (validation order).</param>
/// <param name="ListReportKey">List print-out of the tab (null = none).</param>
public record CodeTableDefinition(string Key, string LegacyForm, string Tab, int Access, string Table, string AuditUserColumn, string AuditDateColumn,
    IReadOnlyList<CodeField> Fields, string? ListReportKey)
{
    /// <summary>The key field.</summary>
    public CodeField KeyField => Fields.Single(f => f.IsKey);
}

/// <summary>
/// The simple code tables of the maintenance forms, transcribed from their Save / Search / Delete routines (frmCheckingItems,
/// frmMaintenanceCal, frmMaintenanceOthers, frmMaintenanceBNFDental, frmMaintenanceBenefits). Desktop behaviour kept: duplicate check
/// on the code only, Edit updates the first row with the code, Delete removes every row with the code.
/// </summary>
public static class CodeTableCatalog
{
    private const string Desc = "Please Enter Description";

    /// <summary>All code tables.</summary>
    public static readonly IReadOnlyList<CodeTableDefinition> All =
    [
        new("check-items", "frmCheckingItems", "Checking Items", AccessRight.MaintenanceCheckItem, "CheckItems", "ChkLastUpdateUser", "ChkLastUpdateDate",
        [
            new("CatCode", "Category Code", 3, "Please Enter Category Code", 3, "Category Code must be 3 integer!", Search: SearchMatch.EqualsFirst3, IsKey: true),
            new("CatName", "Category Name", 30, "Please Enter Category Description"),
            new("Amount", "Amount", 8, "Please Enter Amount", Kind: FieldKind.Number),
        ], "check-items"),

        new("co-payment", "frmMaintenanceCal", "Co-Payment Cal.", AccessRight.MaintenanceCalculation, "CoPayment", "COPLastUpdateUser", "COPLastUpdateDate",
        [
            new("COPCode", "Co-Payment Code", 3, "Please Enter Code", 3, "Co-Payment Code must be 3 integer!", Search: SearchMatch.EqualsFirst3, IsKey: true),
            new("COPDesc", "Description", 100, Desc),
            new("COPRemark", "Remarks", 100),
        ], "co-payment"),
        // The desktop saved FEE's audit columns into COPLastUpdateUser / COPLastUpdateDate, which FEE does not have (silently skipped
        // by On Error Resume Next); the web writes FEE's own columns (technical fix, no business effect).
        new("fee", "frmMaintenanceCal", "Fee Payable 2 Clinic", AccessRight.MaintenanceCalculation, "FEE", "FeeLastUpdateUser", "FeeLastUpdateDate",
        [
            new("FeeCode", "Fee Code", 3, "Please Enter Fee Code", 3, "Fee Code must be 3 integer!", Search: SearchMatch.EqualsFirst3, IsKey: true),
            new("FeeDesc", "Description", 100, Desc),
            new("FeeRemark", "Remarks", 100),
        ], "fee"),

        new("treatment-reason", "frmMaintenanceOthers", "Treatment Reason", AccessRight.MaintenanceCase, "TreatmentReason", "ReasonLastUpdateUser", "ReasonLastUpdateDate",
        [
            new("ReasonCode", "Reason Code", 3, "Please Enter Reason Code", 3, "Reason Code must be 3 integer!", Search: SearchMatch.EqualsFirst3, IsKey: true),
            new("ReasonDesc", "Description", 100, Desc),
            new("ReasonRemark", "Remarks", 200),
        ], "reason"),
        new("procedures", "frmMaintenanceOthers", "Procedures", AccessRight.MaintenanceCase, "Procedures", "ProLastUpdateUser", "ProLastUpdateDate",
        [
            new("ProcedureCode", "Procedure Code", 4, "Please Enter Procedure Code", Search: SearchMatch.EqualsFirst3, IsKey: true),
            new("ProcedureDesc", "Description", 100, Desc),
        ], "procedures"),
        new("drugs", "frmMaintenanceOthers", "Drugs", AccessRight.MaintenanceCase, "Drug", "DrugLastUpdateUser", "DrugLastUpdateDate",
        [
            new("DrugCode", "Drug Code", 3, "Please Enter Drug Code", Search: SearchMatch.EqualsFirst3, IsKey: true),
            new("DrugUnit", "Drug Unit", 10, "Please Enter Drug Unit"),
            new("DrugValue", "Drug Value", 8, "Please Enter Drug Value", Kind: FieldKind.MoneyText),
            new("DrugDesc", "Description", 200),
        ], "drug"),
        new("diagnosis-not-payable", "frmMaintenanceOthers", "Diagnosis Not Payable", AccessRight.MaintenanceCase, "DiagNotPayable", "LastUpdateUser", "LastUpdateDate",
        [
            new("PLNCode", "Plan Code", 3, "Please Enter Plan Code", Search: SearchMatch.EqualsValue, IsKey: true),
            new("DiagCode", "Diagnosis Code", 3, "Please Enter Diagnosis Code", Search: SearchMatch.EqualsValue),
        ], "diagnosis-not-payable"),
        new("repudiation", "frmMaintenanceOthers", "Repudiation", AccessRight.MaintenanceCase, "Repudiation", "RepLastUpdateUser", "RepLastUpdateDate",
        [
            new("RepCode", "Repudiation Code", 4, "Please Enter Repudiation Code", Search: SearchMatch.EqualsValue, IsKey: true),
            new("RepDesc", "Description", 500, Desc),
        ], "repudiation"),

        new("tooth-category", "frmMaintenanceBNFDental", "Tooth Category", AccessRight.MaintenanceBenefitMember, "ToothRefCategory", "ToothRefCatLastUpdateUser", "ToothRefCatLastUpdateDate",
        [
            new("ToothRefCatCode", "Category Code", 2, "Please Enter Category Code", Search: SearchMatch.EqualsValue, IsKey: true),
            new("ToothRefCatDesc", "Description", 50, Desc, Search: SearchMatch.Contains),
        ], "tooth-category"),
        new("tooth-reference", "frmMaintenanceBNFDental", "Tooth Reference", AccessRight.MaintenanceBenefitMember, "ToothRef", "ToothLastUpdateUser", "ToothLastUpdateDate",
        [
            new("ToothRefCode", "Code", 3, "Please Enter Code", Search: SearchMatch.EqualsValue, IsKey: true),
            new("ToothRefCat", "Category", 1, "Please Enter Category", Kind: FieldKind.ChoiceFirstChar, Search: SearchMatch.EqualsFirstChar)
            {
                Choices = new ChoiceSource("ToothRefCategory", "ToothRefCatCode", "ToothRefCatDesc"),
            },
            new("ToothRefDesc", "Description", 255, Desc, Search: SearchMatch.Contains),
        ], "tooth-reference"),
        new("dental-benefit-category", "frmMaintenanceBNFDental", "Benefit Category", AccessRight.MaintenanceBenefitMember, "DentalBNFCategory", "DentalBNFCATLastUpdateUser", "DentalBNFCATLastUpdateDate",
        [
            new("DentalBNFCATCode", "Category Code", 10, "Please Enter Category Code", Search: SearchMatch.EqualsValue, IsKey: true),
            new("DentalBNFCATDesc", "Description", 50, Desc, Search: SearchMatch.Contains),
        ], "dental-benefit-category"),

        new("benefit-category", "frmMaintenanceBenefits", "Category", AccessRight.MaintenanceBenefitMember, "BNFCategoryNew", "BNFCATLastUpdateUser", "BNFCATLastUpdateDate",
        [
            new("BNFCATCode", "Category Code", 3, "Please Enter Code", 3, "Category Code must be 3 integer!", Search: SearchMatch.EqualsValue, IsKey: true),
            new("BNFCATDescription", "Description", 255, Desc),
        ], "benefit-category"),
    ];

    /// <summary>The code table with <paramref name="key"/>, or null.</summary>
    public static CodeTableDefinition? Find(string key) => All.FirstOrDefault(t => string.Equals(t.Key, key, StringComparison.OrdinalIgnoreCase));
}
