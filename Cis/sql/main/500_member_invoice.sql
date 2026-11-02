-- Invoice > Clinic > Member (frmInvoiceMBM): reimbursement invoices. Go reuses the clinic invoice procedures; the desktop saves with recordset
-- inserts (not the clinic's procedures), reproduced here as parameterised INSERTs.

-- ValidateIbmUser: sp_IBMMemberInformation → Number ("0" = not an IBM member).
CREATE OR ALTER PROCEDURE cis.InvMember_Ibm @mbmNumber nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_IBMMemberInformation @MemeberNo = @mbmNumber;
END
GO

-- SaveCase: CISCAS.
CREATE OR ALTER PROCEDURE cis.InvMember_Case
    @casNumber nvarchar(20), @mbmNumber nvarchar(50), @coverId nvarchar(2), @userCode nvarchar(50), @today datetime, @clinicCode nvarchar(10),
    @npClinic nvarchar(200), @invoiceNo nvarchar(50), @invoiceDate datetime, @invoiceStmtAmt money, @dateVisit datetime, @mcDays money,
    @visitTime nvarchar(10) = NULL, @opBatchNo nvarchar(30),
    @t1 nvarchar(200), @t2 nvarchar(200), @t3 nvarchar(200), @t4 nvarchar(200), @t5 nvarchar(200), @t6 nvarchar(200), @t7 nvarchar(200),
    @d1 nvarchar(10), @d2 nvarchar(10), @d3 nvarchar(10), @d4 nvarchar(10), @d5 nvarchar(10), @d6 nvarchar(10), @d7 nvarchar(10),
    @q1 nvarchar(10), @q2 nvarchar(10), @q3 nvarchar(10), @q4 nvarchar(10), @q5 nvarchar(10), @q6 nvarchar(10), @q7 nvarchar(10),
    @receiptDate datetime, @receiptMode nvarchar(10), @remark nvarchar(500), @r1 nvarchar(20), @r2 nvarchar(20), @r3 nvarchar(20), @r4 nvarchar(20),
    @r5 nvarchar(20), @diagnosis5 nvarchar(200), @mcStatus int, @noOfDiag int, @noOfMedication int, @clnType nvarchar(2) = NULL,
    @repReason nvarchar(500), @repAction nvarchar(500), @repAmt nvarchar(20), @isPharmacy nvarchar(2)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.CISCAS (CASNumber, MBMNumber, CoverID, CASStatus, Claimability, CaseType, Pending, RegisterDate, RegisterBy, LastUpdateDate,
        LastUpdateBy, ClinicCode, NPClinic, InvoiceNo, InvoiceDate, InvoiceStmtAmt, DateVisit, MCDays, VisitTime, OPBatchNo,
        MedInjType1, MedInjType2, MedInjType3, MedInjType4, MedInjType5, MedInjType6, MedInjType7, MedInjDays1, MedInjDays2, MedInjDays3, MedInjDays4,
        MedInjDays5, MedInjDays6, MedInjDays7, MedInjQty1, MedInjQty2, MedInjQty3, MedInjQty4, MedInjQty5, MedInjQty6, MedInjQty7, InvoiceReceiptDate,
        InvoiceReceiptMode, CasRemark, CASTreatReason1, CASTreatReason2, CASTreatReason3, CASTreatReason4, CASTreatReason5, CASDiagnosis5, MCStatus,
        NoOfDiag, NoOfMedication, CLNType, RepudiationReason, CASRepAction, CASRepAmt, IsPharmacy)
    VALUES (@casNumber, @mbmNumber, @coverId, 'O', 'A', 'C', 'N', @today, @userCode, @today, @userCode, @clinicCode, @npClinic, @invoiceNo, @invoiceDate,
        @invoiceStmtAmt, @dateVisit, @mcDays, @visitTime, @opBatchNo, @t1, @t2, @t3, @t4, @t5, @t6, @t7, @d1, @d2, @d3, @d4, @d5, @d6, @d7,
        @q1, @q2, @q3, @q4, @q5, @q6, @q7, @receiptDate, @receiptMode, @remark, @r1, @r2, @r3, @r4, @r5, @diagnosis5, @mcStatus, @noOfDiag,
        @noOfMedication, @clnType, @repReason, @repAction, @repAmt, @isPharmacy);
END
GO

-- SaveInvoice: ClaimInvoice (amount less the co-pay), paid to the member.
CREATE OR ALTER PROCEDURE cis.InvMember_Invoice @casNumber nvarchar(20), @userCode nvarchar(50), @today datetime, @billNo nvarchar(50), @billDate datetime,
    @amount money
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Ws_ClaimInvoice (CASNumber, INVEnteredDate, INVEnteredBy, INVLastUpdateBy, INVLastUpdateDate, BillNo, BillDate, INVAmount, INVBillAmt,
        INVPayTo, INVPayee)
    VALUES (@casNumber, @today, @userCode, @userCode, @today, @billNo, @billDate, @amount, @amount, 'M', 'MEMBER');
END
GO

-- saveActual: ClaimInvoiceActualNew (the 28 items, co-payment, descriptions, vision, others without deduction).
CREATE OR ALTER PROCEDURE cis.InvMember_Actual @casNumber nvarchar(20),
    @a1 money, @a2 money, @a3 money, @a4 money, @a5 money, @a6 money, @a7 money, @a8 money, @a9 money, @a10 money, @a11 money, @a12 money, @a13 money,
    @a14 money, @a15 money, @a16 money, @a17 money, @a18 money, @a19 money, @a20 money, @a21 money, @a22 money, @a23 money, @a24 money, @a25 money,
    @a26 money, @a27 money, @a28 money, @coPayment money, @pOthDesc nvarchar(200), @lOthDesc nvarchar(200), @iOthDesc nvarchar(200), @vision money,
    @othersNoDed money
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Ws_ClaimInvoiceActualNew (CASNumber, CLMGPFeeBasicMed, CLMGPFeeLTMed, CLMGPFeeMedInj, CLMSPFeeBasicMed, CLMSPFeeLTMed, CLMSPFeeMedInj,
        CLMDental, CLMMaternity, CLMPRVCheckup, CLMXRay, CLMUltraSound, CLMECG, CLMDressing, CLMNebulizer, CLMID, CLMTSSTO, CLMMinorSurgery, CLMPOthers,
        CLMBloodDrawing, CLMBloodGlucose, CLMUrine, CLMPapSmear, CLMLOthers, CLMDTPPolio, CLMDTPPolioHIB, CLMHepatitisB, CLMMMR, CLMIOthers, CLMCoPayment,
        CLMPOthersDesc, CLMLOthersDesc, CLMIOthersDesc, CLMVision, CLMOthersNoDed)
    VALUES (@casNumber, @a1, @a2, @a3, @a4, @a5, @a6, @a7, @a8, @a9, @a10, @a11, @a12, @a13, @a14, @a15, @a16, @a17, @a18, @a19, @a20, @a21, @a22, @a23,
        @a24, @a25, @a26, @a27, @a28, @coPayment, @pOthDesc, @lOthDesc, @iOthDesc, @vision, @othersNoDed);
END
GO

-- saveActualOthers: ClaimInvoiceActualOthers (GP / SP fee parts).
CREATE OR ALTER PROCEDURE cis.InvMember_ActualOthers @casNumber nvarchar(20), @gpFee money, @gpLtMed money, @gpInj money, @gpMed money, @spFee money,
    @spLtMed money, @spInj money, @spMed money
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Ws_ClaimInvoiceActualOthers (CASNumber, GPFee, GPLTMed, GPInj, GPMed, SPFee, SPLTMed, SPInj, SPMed)
    VALUES (@casNumber, @gpFee, @gpLtMed, @gpInj, @gpMed, @spFee, @spLtMed, @spInj, @spMed);
END
GO

-- SaveINVClaim: Claim.
CREATE OR ALTER PROCEDURE cis.InvMember_Claim @casNumber nvarchar(20), @userCode nvarchar(50), @today datetime, @payableM money, @clnExcess money,
    @totalActual money, @finalApprove money, @balanceLimit money, @previousBalance money, @paidM money, @paidC money, @postageFee money, @adjAnnLmt money,
    @nonGpFee money, @bordxAmt money, @coPay money, @memo char(1), @refLetter nvarchar(2), @deptName nvarchar(200), @costCenter nvarchar(100),
    @deptCode nvarchar(100), @subsidiary nvarchar(100), @grpCoCode nvarchar(50)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Ws_Claim (CASNumber, ClaimOpenDate, ClaimOpenBy, ClaimCloseDate, ClaimClosedBy, ClaimLastUpdateDate, ClaimLastUpdateUser,
        CLMPayableByMedix2M, CLMCLNExcess, CLMTotalActual, CLMFinalApprove, CLMBalanceLimit, CLMPreviousBalance, CLMPaidBYMedix2M, CLMPaidBYMedix2C,
        CLMPostageFee, CLMAdjAnnLmt, ClmNonGPFee, CLMBordxAmt, CLMCoPayColl, CLMCoPayCLNColl, CLMBordxMemo, CLMRefLetter, CLMDeptName, CLMCostCenter,
        CLMDeptCode, CLMSubsidiary, CLMGrpComCode)
    VALUES (@casNumber, @today, @userCode, @today, @userCode, @today, @userCode, @payableM, @clnExcess, @totalActual, @finalApprove, @balanceLimit,
        @previousBalance, @paidM, @paidC, @postageFee, @adjAnnLmt, @nonGpFee, @bordxAmt, @coPay, @coPay, @memo, @refLetter, @deptName, @costCenter,
        @deptCode, @subsidiary, @grpCoCode);
END
GO

-- UpdateMBMBalance (member invoice): adjusted annual limit, accumulated amount, balance and the health limit.
CREATE OR ALTER PROCEDURE cis.InvMember_Balance @mbmNumber nvarchar(50), @limitCover nvarchar(2) = NULL, @adjAnnLmt money, @approved money,
    @deduct money, @full bit, @health bit
AS
BEGIN
    SET NOCOUNT ON;
    IF @limitCover IS NULL
        UPDATE cis.Hdb_MBMCLNOthers SET MBMAdjAnnLmt = COALESCE(MBMAdjAnnLmt, 0) + @adjAnnLmt, MBMAccumAmt = COALESCE(MBMAccumAmt, 0) + @approved,
            MBMCLNBalLmt = CASE WHEN @full = 1 THEN MBMCLNBalLmt WHEN COALESCE(MBMCLNBalLmt, 0) - @deduct < 0 THEN 0 ELSE COALESCE(MBMCLNBalLmt, 0) - @deduct END,
            MBMHealthlimit = CASE WHEN @health = 1 THEN COALESCE(TRY_CONVERT(money, MBMHealthlimit), 0) - @approved ELSE MBMHealthlimit END
        WHERE MBMNumber = @mbmNumber;
    ELSE
        UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCAdjAnnLmt = COALESCE(TRY_CONVERT(money, MBMCAdjAnnLmt), 0) + @adjAnnLmt,
            MBMCAccumAmt = COALESCE(TRY_CONVERT(money, MBMCAccumAmt), 0) + @approved,
            MBMCCLNBalLmt = CASE WHEN @full = 1 THEN MBMCCLNBalLmt
                                 WHEN COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) - @deduct < 0 THEN 0 ELSE COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) - @deduct END,
            MBMCHealthlimit = CASE WHEN @health = 1 THEN COALESCE(TRY_CONVERT(money, MBMCHealthlimit), 0) - @approved ELSE MBMCHealthlimit END
        WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @limitCover;
END
GO
