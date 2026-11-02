-- Medix.bas MemberSuppInfo: the member, plan and limit details the invoice screens and the member info pop-up (FrmINVMBMinfo) show.
-- Result sets: 1 member (VIEW_MBM_Clinical), 2 group company, 3 MBMCLNOthers, 4 PLN, 5 CoPaymentMBM, 6 AnnualLimit.SuppLimitStatus,
-- 7 covered person, 8 MBMCoveredPersonsCLN of the cover and of cover 01 (the shared-limit cover). The rules are in MemberSuppService.
CREATE OR ALTER PROCEDURE cis.MemberSupp_Info @mbmNumber nvarchar(50), @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @group nvarchar(200), @hlt nvarchar(20), @pln nvarchar(20), @ins nvarchar(20), @pay nvarchar(20);
    SELECT TOP (1) @group = MBMGroupCompany, @hlt = RTRIM(HLTCode) FROM cis.Hdb_VIEW_MBM_Clinical WHERE MBMNumber = @mbmNumber AND MBMCLNDataStatus <> 'D';
    SELECT TOP (1) @pln = RTRIM(PLNType), @ins = RTRIM(INSType), @pay = RTRIM(PAYCode) FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber = @mbmNumber;

    SELECT TOP (1) MBMNumber, MBMName, MBMIcBcPp, PAYCode, MBMCLNStatus, MBMGroupCompany, MBMDOB, MBMCLNDataStatus, RTRIM(HLTCode) AS HLTCode,
           MBMDepartment, MBMCLNMCDaysBal
    FROM cis.Hdb_VIEW_MBM_Clinical WHERE MBMNumber = @mbmNumber AND MBMCLNDataStatus <> 'D';
    SELECT TOP (1) GRPMCOStatus, GRPMCStatus, GRPPostageFees, GRPBordxFees FROM dbo.GRPCompany WHERE GrpCoName = @group;
    SELECT TOP (1) MBMPolNumber, MBMPolEffDate, MBMPolExpDate, MBMCLNBalLmt, MBMCLNVisitLmt, MBMAccumAmt, MBMCardType, RTRIM(INSType) AS INSType,
           RTRIM(PLNType) AS PLNType, MBMCLNLTM, RTRIM(PAYCode) AS PAYCode, MBMMEPLmt, MBMDentalLmt, MBMSpeclmt, MBMDentalExtLmt, MBMCancelDate,
           MBMFirstSpecLmt
    FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber = @mbmNumber;
    SELECT TOP (1) AnnLmtIND, VisLmtIND, ClnContract, PDentalStatus, SDentalStatus, PSpecStatus, SSpecStatus, PMEPStatus, SMEPStatus, BNFStatus,
           PDentalLimit, SDentalLimit, PSpecLimit, SSpecLimit, ANNLmtIndSupp, VISLmtIndSupp, GrossAmtStatus, PLNMMIndPerc, ASOInd
    FROM dbo.PLN WHERE PLNCode = @pln AND HLTCode = @hlt AND PAYCode = @pay;
    SELECT TOP (1) CoAmount FROM dbo.CoPaymentMBM WHERE HLTType = @hlt AND PlnCode = @pln;
    SELECT TOP (1) SuppLimitStatus FROM dbo.AnnualLimit WHERE PLNCode = @pln AND HLTCode = @hlt AND INSCode = @ins;
    SELECT TOP (1) MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB, LTRIM(RTRIM(MBMCRELCode)) AS MBMCRELCode
    FROM cis.Hdb_MBMCoveredPersons WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @covId;
    SELECT MBMCCoverID, MBMCCLNBalLmt, MBMCCLNVisitLmt, MBMCCLNDentalLmt, MBMCCLNSpecLmt, MBMCCLNMEPLmt, MBMCCLNDentalExtLmt, MBMCCLNFirstSpecLmt,
           MBMCAccumAmt, MBMCCardType, MBMCCancelDate, MBMCCLNExpDate, MBMCCLNEffDate, MBMCCLNStatus
    FROM cis.Hdb_MBMCoveredPersonsCLN WHERE MBMCNumber = @mbmNumber AND MBMCCoverID IN (@covId, '01');
END
GO
