-- Invoice > General Exclusions (Enquiry) (frmGeneralBenefitExcEnq): CISDB GeneralExclusions per group company (read-only; the desktop's
-- Update button is hidden). frmGeneralBenefitExclusion (the editing form) is not in the source.

-- ComboList: group company names.
CREATE OR ALTER PROCEDURE cis.GeneralExclusions_GroupCompanies AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
END
GO

-- CallExclusions: the group company's exclusion flags and remarks.
CREATE OR ALTER PROCEDURE cis.GeneralExclusions_Get @groupCompany nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) Dermatitis, Gynaecological, Dental, Optical, MedCheckup, PreEmpCheckup, Hypertension, Diabetes, IHDCAD, Thyroid, Immunisation, HRT,
           PregnacyTest, Antenatal, Abortion, Infertility, Contraception, Genitourinary, Vitamins, BloodTest, XRAY, PapSmear, Typhoid, LTMedication, Remarks
    FROM dbo.GeneralExclusions WHERE GRPCompany = @groupCompany;
END
GO
