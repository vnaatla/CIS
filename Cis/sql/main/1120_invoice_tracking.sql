-- Tracking > Invoice (FrmInvoiceTracking): CISWorksheet ClaimBordxModule case 25 (VIEW_Invoice_Tracking_2020, or VIEW_Invoice_Tracking_2020_BTBG when a bordx no.
-- criterion is used) with the pasted criteria replaced by parameters. The "blank" switches (@...Blank = 1) are the desktop's "empty" check boxes.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.InvTrk_Search
    @caseFrom nvarchar(20) = NULL, @caseTo nvarchar(20) = NULL, @caseBlank bit = 0,
    @payTo nvarchar(1) = NULL, @payToBlank bit = 0, @clinic nvarchar(10) = NULL, @payeeBlank bit = 0,
    @payor nvarchar(10) = NULL, @payorBlank bit = 0, @company nvarchar(200) = NULL, @companyBlank bit = 0,
    @patient nvarchar(200) = NULL, @patientBlank bit = 0, @mbmFrom nvarchar(30) = NULL, @mbmTo nvarchar(30) = NULL, @mbmBlank bit = 0,
    @billFrom nvarchar(50) = NULL, @billTo nvarchar(50) = NULL, @billBlank bit = 0,
    @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL, @bordxBlank bit = 0,
    @statuses nvarchar(10) = NULL, @pending bit = 0, @claimability nvarchar(1) = NULL,
    @planFrom nvarchar(20) = NULL, @planTo nvarchar(20) = NULL, @planBlank bit = 0,
    @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL,
    @icFrom nvarchar(30) = NULL, @icTo nvarchar(30) = NULL, @icBlank bit = 0,
    @visitFrom datetime = NULL, @visitTo datetime = NULL, @visitBlank bit = 0,
    @regFrom date = NULL, @regTo date = NULL, @regBlank bit = 0,
    @billDateFrom datetime = NULL, @billDateTo datetime = NULL, @billDateBlank bit = 0,
    @enteredFrom datetime = NULL, @enteredTo datetime = NULL, @enteredBlank bit = 0,
    @amountFrom money = NULL, @amountTo money = NULL, @amountBlank bit = 0,
    @accMonthFrom int = NULL, @accMonthTo int = NULL, @accMonthBlank bit = 0, @accYearFrom int = NULL, @accYearTo int = NULL, @accYearBlank bit = 0,
    @clinicType nvarchar(1) = NULL, @opBatchFrom nvarchar(30) = NULL, @opBatchTo nvarchar(30) = NULL, @opBatchDateFrom datetime = NULL,
    @opBatchDateTo datetime = NULL, @opBatchBlank bit = 0, @ffs nvarchar(1) = NULL, @omni bit = 0, @tims bit = 0,
    @subsidiary nvarchar(200) = NULL, @department nvarchar(200) = NULL,
    @omniRefFrom nvarchar(30) = NULL, @omniRefTo nvarchar(30) = NULL, @omniRefBlank bit = 0, @omniSent nvarchar(1) = NULL,
    @sj bit = 0, @flex int = NULL, @paid nvarchar(1) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @btbg bit = CASE WHEN @bordxFrom IS NOT NULL OR @bordxTo IS NOT NULL OR @bordxBlank = 1 THEN 1 ELSE 0 END;
    DECLARE @sql nvarchar(max) = CAST(N'' AS nvarchar(max)) + N'SELECT * FROM CISWorksheet.dbo.' + CASE @btbg WHEN 1 THEN N'VIEW_Invoice_Tracking_2020_BTBG' ELSE N'VIEW_Invoice_Tracking_2020' END + N'
    WHERE (@caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@caseTo IS NULL OR CASNumber <= @caseTo) AND (@caseBlank = 0 OR CASNumber IS NULL OR CASNumber = '''')
      AND (@payTo IS NULL OR INVPayTo = @payTo) AND (@payToBlank = 0 OR INVPayTo IS NULL OR INVPayTo = '''')
      AND (@clinic IS NULL OR ClinicCode = @clinic) AND (@payeeBlank = 0 OR INVPayee IS NULL OR INVPayee = '''')
      AND (@payor IS NULL OR PAYCode = @payor) AND (@payorBlank = 0 OR PAYCode IS NULL OR PAYCode = '''')
      AND (@company IS NULL OR MBMGroupCompany = @company) AND (@companyBlank = 0 OR MBMGroupCompany IS NULL OR MBMGroupCompany = '''')
      AND (@patient IS NULL OR PatientName = @patient) AND (@patientBlank = 0 OR PatientName IS NULL OR PatientName = '''')
      AND (@mbmFrom IS NULL OR MBMNumber >= @mbmFrom) AND (@mbmTo IS NULL OR MBMNumber <= @mbmTo) AND (@mbmBlank = 0 OR MBMNumber IS NULL OR MBMNumber = '''')
      AND (@billFrom IS NULL OR BillNo >= @billFrom) AND (@billTo IS NULL OR BillNo <= @billTo) AND (@billBlank = 0 OR BillNo IS NULL OR BillNo = '''')
      AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
      AND (@bordxBlank = 0 OR BordxRefNo IS NULL OR BordxRefNo = '''')
      AND (@statuses IS NULL OR CHARINDEX(RTRIM(CASStatus), @statuses) > 0) AND (@pending = 0 OR Pending = ''Y'')
      AND (@claimability IS NULL OR Claimability = @claimability)
      AND (@planFrom IS NULL OR PLNCode >= @planFrom) AND (@planTo IS NULL OR PLNCode <= @planTo) AND (@planBlank = 0 OR PLNCode IS NULL OR PLNCode = '''')
      AND (@bordxDateFrom IS NULL OR BordxDate >= @bordxDateFrom) AND (@bordxDateTo IS NULL OR BordxDate <= @bordxDateTo)
      AND (@icFrom IS NULL OR ICNo >= @icFrom) AND (@icTo IS NULL OR ICNo <= @icTo) AND (@icBlank = 0 OR ICNo IS NULL)
      AND (@visitFrom IS NULL OR DateVisit >= @visitFrom) AND (@visitTo IS NULL OR DateVisit <= @visitTo) AND (@visitBlank = 0 OR DateVisit IS NULL OR DateVisit = '''')
      AND (@regFrom IS NULL OR CAST(RegisterDate AS date) >= @regFrom) AND (@regTo IS NULL OR CAST(RegisterDate AS date) <= @regTo)
      AND (@regBlank = 0 OR RegisterDate IS NULL OR RegisterDate = '''')
      AND (@billDateFrom IS NULL OR BillDate >= @billDateFrom) AND (@billDateTo IS NULL OR BillDate <= @billDateTo) AND (@billDateBlank = 0 OR BillDate IS NULL OR BillDate = '''')
      AND (@enteredFrom IS NULL OR INVEnteredDate >= @enteredFrom) AND (@enteredTo IS NULL OR INVEnteredDate <= @enteredTo)
      AND (@enteredBlank = 0 OR INVEnteredDate IS NULL OR INVEnteredDate = '''')
      AND (@amountFrom IS NULL OR INVAmount >= @amountFrom) AND (@amountTo IS NULL OR INVAmount <= @amountTo) AND (@amountBlank = 0 OR INVAmount IS NULL)
      AND (@accMonthFrom IS NULL OR TRY_CAST(AccPeriodMonth AS int) >= @accMonthFrom) AND (@accMonthTo IS NULL OR TRY_CAST(AccPeriodMonth AS int) <= @accMonthTo)
      AND (@accMonthBlank = 0 OR AccPeriodMonth IS NULL)
      AND (@accYearFrom IS NULL OR TRY_CAST(AccPeriodYear AS int) >= @accYearFrom) AND (@accYearTo IS NULL OR TRY_CAST(AccPeriodYear AS int) <= @accYearTo)
      AND (@accYearBlank = 0 OR AccPeriodYear IS NULL)
      AND (@clinicType IS NULL OR (@clinicType = ''0'' AND CLNType = ''S'') OR (@clinicType = ''1'' AND CLNType = ''G'')
           OR (@clinicType = ''2'' AND (CLNType = '''' OR CLNType IS NULL)))
      AND (@opBatchFrom IS NULL OR OPBatchNo >= @opBatchFrom) AND (@opBatchTo IS NULL OR OPBatchNo <= @opBatchTo)
      AND (@opBatchDateFrom IS NULL OR OPBatchDate >= @opBatchDateFrom) AND (@opBatchDateTo IS NULL OR OPBatchDate <= @opBatchDateTo)
      AND (@opBatchBlank = 0 OR OPBatchNo IS NULL OR OPBatchNo = '''')
      AND (@ffs IS NULL OR (@ffs = ''Y'' AND CLNFFSStatus = ''Y'') OR (@ffs = ''N'' AND (CLNFFSStatus = ''N'' OR CLNFFSStatus = '''' OR CLNFFSStatus IS NULL)))
      AND (@omni = 0 OR GRPOMNICLM = ''Y'') AND (@tims = 0 OR GRPTIMS = ''Y'')
      AND (@subsidiary IS NULL OR clmsubsidiary = @subsidiary) AND (@department IS NULL OR CLMDeptName = @department)
      AND (@omniRefBlank = 0 OR OMNIRefNo IS NULL OR OMNIRefNo = '''') AND (@omniRefFrom IS NULL OR OMNIRefNo >= @omniRefFrom) AND (@omniRefTo IS NULL OR OMNIRefNo <= @omniRefTo)
      AND (@omniSent IS NULL OR (@omniSent = ''Y'' AND CLMOMNISent = ''Y'') OR (@omniSent = ''N'' AND CLMOMNISent IS NULL))
      AND ((@sj = 0 AND SUBSTRING(CASNumber, 1, 2) <> ''SJ'') OR (@sj = 1 AND SUBSTRING(CASNumber, 1, 2) = ''SJ''))
      AND (@flex IS NULL OR (@flex = 0 AND (FBEmployee IS NULL OR (FBFamily = 1 AND CoverID <> ''00''))) OR (@flex = 1 AND FBEmployee IS NOT NULL AND CoverID = ''00''))
      AND (@paid IS NULL OR @payTo IS NULL OR (@paid = ''Y'' AND @payTo = ''C'' AND PVCLNNo IS NOT NULL) OR (@paid = ''Y'' AND @payTo = ''M'' AND PVNo IS NOT NULL)
           OR (@paid = ''N'' AND @payTo = ''C'' AND PVCLNNo IS NULL) OR (@paid = ''N'' AND @payTo = ''M'' AND PVNo IS NULL))
    OPTION (RECOMPILE)';
    EXEC sp_executesql @sql, N'@caseFrom nvarchar(20), @caseTo nvarchar(20), @caseBlank bit, @payTo nvarchar(1), @payToBlank bit, @clinic nvarchar(10), @payeeBlank bit,
        @payor nvarchar(10), @payorBlank bit, @company nvarchar(200), @companyBlank bit, @patient nvarchar(200), @patientBlank bit, @mbmFrom nvarchar(30), @mbmTo nvarchar(30),
        @mbmBlank bit, @billFrom nvarchar(50), @billTo nvarchar(50), @billBlank bit, @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @bordxBlank bit, @statuses nvarchar(10),
        @pending bit, @claimability nvarchar(1), @planFrom nvarchar(20), @planTo nvarchar(20), @planBlank bit, @bordxDateFrom datetime, @bordxDateTo datetime,
        @icFrom nvarchar(30), @icTo nvarchar(30), @icBlank bit, @visitFrom datetime, @visitTo datetime, @visitBlank bit, @regFrom date, @regTo date, @regBlank bit,
        @billDateFrom datetime, @billDateTo datetime, @billDateBlank bit, @enteredFrom datetime, @enteredTo datetime, @enteredBlank bit, @amountFrom money, @amountTo money,
        @amountBlank bit, @accMonthFrom int, @accMonthTo int, @accMonthBlank bit, @accYearFrom int, @accYearTo int, @accYearBlank bit, @clinicType nvarchar(1),
        @opBatchFrom nvarchar(30), @opBatchTo nvarchar(30), @opBatchDateFrom datetime, @opBatchDateTo datetime, @opBatchBlank bit, @ffs nvarchar(1), @omni bit, @tims bit,
        @subsidiary nvarchar(200), @department nvarchar(200), @omniRefFrom nvarchar(30), @omniRefTo nvarchar(30), @omniRefBlank bit, @omniSent nvarchar(1), @sj bit,
        @flex int, @paid nvarchar(1)',
        @caseFrom, @caseTo, @caseBlank, @payTo, @payToBlank, @clinic, @payeeBlank, @payor, @payorBlank, @company, @companyBlank, @patient, @patientBlank, @mbmFrom, @mbmTo,
        @mbmBlank, @billFrom, @billTo, @billBlank, @bordxFrom, @bordxTo, @bordxBlank, @statuses, @pending, @claimability, @planFrom, @planTo, @planBlank, @bordxDateFrom,
        @bordxDateTo, @icFrom, @icTo, @icBlank, @visitFrom, @visitTo, @visitBlank, @regFrom, @regTo, @regBlank, @billDateFrom, @billDateTo, @billDateBlank, @enteredFrom,
        @enteredTo, @enteredBlank, @amountFrom, @amountTo, @amountBlank, @accMonthFrom, @accMonthTo, @accMonthBlank, @accYearFrom, @accYearTo, @accYearBlank, @clinicType,
        @opBatchFrom, @opBatchTo, @opBatchDateFrom, @opBatchDateTo, @opBatchBlank, @ffs, @omni, @tims, @subsidiary, @department, @omniRefFrom, @omniRefTo, @omniRefBlank,
        @omniSent, @sj, @flex, @paid;
END
GO

-- Tracking (Jupiter) (FrmInvoiceTracking2): ClaimBordxModule case 1 (VIEW_Invoice_Tracking_2016, or _2016_BTBG with a bordx criterion; MSNO = the row number
-- within the visit month). Its subset of the filters: the clinic compared with INVPayee, one case status, the register date without the day cast.
CREATE OR ALTER PROCEDURE cis.InvTrkJup_Search
    @caseFrom nvarchar(20) = NULL, @caseTo nvarchar(20) = NULL, @caseBlank bit = 0,
    @payTo nvarchar(1) = NULL, @payToBlank bit = 0, @clinic nvarchar(10) = NULL, @payeeBlank bit = 0,
    @payor nvarchar(10) = NULL, @payorBlank bit = 0, @company nvarchar(200) = NULL, @companyBlank bit = 0,
    @patient nvarchar(200) = NULL, @patientBlank bit = 0, @mbmFrom nvarchar(30) = NULL, @mbmTo nvarchar(30) = NULL, @mbmBlank bit = 0,
    @billFrom nvarchar(50) = NULL, @billTo nvarchar(50) = NULL, @billBlank bit = 0,
    @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL, @bordxBlank bit = 0,
    @statuses nvarchar(10) = NULL, @pending bit = 0, @claimability nvarchar(1) = NULL,
    @planFrom nvarchar(20) = NULL, @planTo nvarchar(20) = NULL, @planBlank bit = 0,
    @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL,
    @icFrom nvarchar(30) = NULL, @icTo nvarchar(30) = NULL, @icBlank bit = 0,
    @visitFrom datetime = NULL, @visitTo datetime = NULL, @visitBlank bit = 0,
    @regFrom date = NULL, @regTo date = NULL, @regBlank bit = 0,
    @billDateFrom datetime = NULL, @billDateTo datetime = NULL, @billDateBlank bit = 0,
    @enteredFrom datetime = NULL, @enteredTo datetime = NULL, @enteredBlank bit = 0,
    @amountFrom money = NULL, @amountTo money = NULL, @amountBlank bit = 0,
    @accMonthFrom int = NULL, @accMonthTo int = NULL, @accMonthBlank bit = 0, @accYearFrom int = NULL, @accYearTo int = NULL, @accYearBlank bit = 0,
    @clinicType nvarchar(1) = NULL, @opBatchFrom nvarchar(30) = NULL, @opBatchTo nvarchar(30) = NULL, @opBatchDateFrom datetime = NULL,
    @opBatchDateTo datetime = NULL, @opBatchBlank bit = 0, @ffs nvarchar(1) = NULL, @omni bit = 0, @tims bit = 0,
    @subsidiary nvarchar(200) = NULL, @department nvarchar(200) = NULL,
    @omniRefFrom nvarchar(30) = NULL, @omniRefTo nvarchar(30) = NULL, @omniRefBlank bit = 0, @omniSent nvarchar(1) = NULL,
    @sj bit = 0, @flex int = NULL, @paid nvarchar(1) = NULL, @status nvarchar(1) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @btbg bit = CASE WHEN @bordxFrom IS NOT NULL OR @bordxTo IS NOT NULL OR @bordxBlank = 1 THEN 1 ELSE 0 END;
    DECLARE @sql nvarchar(max) = CAST(N'' AS nvarchar(max)) + N'SELECT ROW_NUMBER() OVER (PARTITION BY MONTH(DateVisit) ORDER BY DateVisit) AS MSNO, * FROM CISWorksheet.dbo.'
        + CASE @btbg WHEN 1 THEN N'VIEW_Invoice_Tracking_2016_BTBG' ELSE N'VIEW_Invoice_Tracking_2016' END + N'
    WHERE (@caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@caseTo IS NULL OR CASNumber <= @caseTo) AND (@caseBlank = 0 OR CASNumber IS NULL OR CASNumber = '''')
      AND (@payTo IS NULL OR INVPayTo = @payTo) AND (@payToBlank = 0 OR INVPayTo IS NULL OR INVPayTo = '''')
      AND (@clinic IS NULL OR INVPayee = @clinic) AND (@payeeBlank = 0 OR INVPayee IS NULL OR INVPayee = '''')
      AND (@payor IS NULL OR PAYCode = @payor) AND (@payorBlank = 0 OR PAYCode IS NULL OR PAYCode = '''')
      AND (@company IS NULL OR MBMGroupCompany = @company) AND (@companyBlank = 0 OR MBMGroupCompany IS NULL OR MBMGroupCompany = '''')
      AND (@patient IS NULL OR PatientName = @patient) AND (@patientBlank = 0 OR PatientName IS NULL OR PatientName = '''')
      AND (@mbmFrom IS NULL OR MBMNumber >= @mbmFrom) AND (@mbmTo IS NULL OR MBMNumber <= @mbmTo) AND (@mbmBlank = 0 OR MBMNumber IS NULL OR MBMNumber = '''')
      AND (@billFrom IS NULL OR BillNo >= @billFrom) AND (@billTo IS NULL OR BillNo <= @billTo) AND (@billBlank = 0 OR BillNo IS NULL OR BillNo = '''')
      AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
      AND (@bordxBlank = 0 OR BordxRefNo IS NULL OR BordxRefNo = '''')
      AND (@status IS NULL OR CASStatus = @status) AND (@claimability IS NULL OR Claimability = @claimability)
      AND (@planFrom IS NULL OR PLNCode >= @planFrom) AND (@planTo IS NULL OR PLNCode <= @planTo) AND (@planBlank = 0 OR PLNCode IS NULL OR PLNCode = '''')
      AND (@icFrom IS NULL OR ICNo >= @icFrom) AND (@icTo IS NULL OR ICNo <= @icTo) AND (@icBlank = 0 OR ICNo IS NULL)
      AND (@visitFrom IS NULL OR DateVisit >= @visitFrom) AND (@visitTo IS NULL OR DateVisit <= @visitTo) AND (@visitBlank = 0 OR DateVisit IS NULL OR DateVisit = '''')
      AND (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate <= @regTo)
      AND (@regBlank = 0 OR RegisterDate IS NULL OR RegisterDate = '''')
      AND (@billDateFrom IS NULL OR BillDate >= @billDateFrom) AND (@billDateTo IS NULL OR BillDate <= @billDateTo) AND (@billDateBlank = 0 OR BillDate IS NULL OR BillDate = '''')
      AND (@enteredFrom IS NULL OR INVEnteredDate >= @enteredFrom) AND (@enteredTo IS NULL OR INVEnteredDate <= @enteredTo)
      AND (@enteredBlank = 0 OR INVEnteredDate IS NULL OR INVEnteredDate = '''')
      AND (@amountFrom IS NULL OR INVAmount >= @amountFrom) AND (@amountTo IS NULL OR INVAmount <= @amountTo) AND (@amountBlank = 0 OR INVAmount IS NULL)
      AND (@accMonthFrom IS NULL OR TRY_CAST(AccPeriodMonth AS int) >= @accMonthFrom) AND (@accMonthTo IS NULL OR TRY_CAST(AccPeriodMonth AS int) <= @accMonthTo)
      AND (@accMonthBlank = 0 OR AccPeriodMonth IS NULL)
      AND (@accYearFrom IS NULL OR TRY_CAST(AccPeriodYear AS int) >= @accYearFrom) AND (@accYearTo IS NULL OR TRY_CAST(AccPeriodYear AS int) <= @accYearTo)
      AND (@accYearBlank = 0 OR AccPeriodYear IS NULL)
      AND (@clinicType IS NULL OR (@clinicType = ''0'' AND CLNType = ''S'') OR (@clinicType = ''1'' AND CLNType = ''G'')
           OR (@clinicType = ''2'' AND (CLNType = '''' OR CLNType IS NULL)))
      AND (@opBatchFrom IS NULL OR OPBatchNo >= @opBatchFrom) AND (@opBatchTo IS NULL OR OPBatchNo <= @opBatchTo)
      AND (@opBatchBlank = 0 OR OPBatchNo IS NULL OR OPBatchNo = '''')
      AND (@ffs IS NULL OR (@ffs = ''Y'' AND CLNFFSStatus = ''Y'') OR (@ffs = ''N'' AND (CLNFFSStatus = ''N'' OR CLNFFSStatus = '''' OR CLNFFSStatus IS NULL)))
    OPTION (RECOMPILE)';
    EXEC sp_executesql @sql, N'@caseFrom nvarchar(20), @caseTo nvarchar(20), @caseBlank bit, @payTo nvarchar(1), @payToBlank bit, @clinic nvarchar(10), @payeeBlank bit,
        @payor nvarchar(10), @payorBlank bit, @company nvarchar(200), @companyBlank bit, @patient nvarchar(200), @patientBlank bit, @mbmFrom nvarchar(30), @mbmTo nvarchar(30),
        @mbmBlank bit, @billFrom nvarchar(50), @billTo nvarchar(50), @billBlank bit, @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @bordxBlank bit, @statuses nvarchar(10),
        @pending bit, @claimability nvarchar(1), @planFrom nvarchar(20), @planTo nvarchar(20), @planBlank bit, @bordxDateFrom datetime, @bordxDateTo datetime,
        @icFrom nvarchar(30), @icTo nvarchar(30), @icBlank bit, @visitFrom datetime, @visitTo datetime, @visitBlank bit, @regFrom date, @regTo date, @regBlank bit,
        @billDateFrom datetime, @billDateTo datetime, @billDateBlank bit, @enteredFrom datetime, @enteredTo datetime, @enteredBlank bit, @amountFrom money, @amountTo money,
        @amountBlank bit, @accMonthFrom int, @accMonthTo int, @accMonthBlank bit, @accYearFrom int, @accYearTo int, @accYearBlank bit, @clinicType nvarchar(1),
        @opBatchFrom nvarchar(30), @opBatchTo nvarchar(30), @opBatchDateFrom datetime, @opBatchDateTo datetime, @opBatchBlank bit, @ffs nvarchar(1), @omni bit, @tims bit,
        @subsidiary nvarchar(200), @department nvarchar(200), @omniRefFrom nvarchar(30), @omniRefTo nvarchar(30), @omniRefBlank bit, @omniSent nvarchar(1), @sj bit,
        @flex int, @paid nvarchar(1), @status nvarchar(1)',
        @caseFrom, @caseTo, @caseBlank, @payTo, @payToBlank, @clinic, @payeeBlank, @payor, @payorBlank, @company, @companyBlank, @patient, @patientBlank, @mbmFrom, @mbmTo,
        @mbmBlank, @billFrom, @billTo, @billBlank, @bordxFrom, @bordxTo, @bordxBlank, @statuses, @pending, @claimability, @planFrom, @planTo, @planBlank, @bordxDateFrom,
        @bordxDateTo, @icFrom, @icTo, @icBlank, @visitFrom, @visitTo, @visitBlank, @regFrom, @regTo, @regBlank, @billDateFrom, @billDateTo, @billDateBlank, @enteredFrom,
        @enteredTo, @enteredBlank, @amountFrom, @amountTo, @amountBlank, @accMonthFrom, @accMonthTo, @accMonthBlank, @accYearFrom, @accYearTo, @accYearBlank, @clinicType,
        @opBatchFrom, @opBatchTo, @opBatchDateFrom, @opBatchDateTo, @opBatchBlank, @ffs, @omni, @tims, @subsidiary, @department, @omniRefFrom, @omniRefTo, @omniRefBlank,
        @omniSent, @sj, @flex, @paid, @status;
END
GO

-- The combos, as the desktop's lookup procedures: @kind clinics (ChkClinic2, SearchCLN), payors (ChkPayor2, SearchPayor), companies (ChkGrpCom2, SearchPayGrp for the
-- payor's first 2 characters), subsidiaries / departments (group company LostFocus, HISDB SearchSubsidiary / SearchDepartment).
CREATE OR ALTER PROCEDURE cis.InvTrk_Lookup @kind nvarchar(20), @key nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @kind = 'clinics' SELECT RTRIM(CLNCode) + ' - ' + RTRIM(CLNName) FROM cis.Hdb_CLN WHERE CLNPanel = 'C' OR CLNPanel = 'Y';
    ELSE IF @kind = 'payors' SELECT RTRIM(PAYCode) + ' - ' + RTRIM(PAYCompanyName) FROM dbo.PAY;
    ELSE IF @kind = 'companies' SELECT RTRIM(GRPCoName) FROM dbo.GRPCompany WHERE GRPCoPAYCode = LEFT(@key, 2);
    ELSE IF @kind = 'subsidiaries' SELECT RTRIM(MBMGRPSubsidiary) FROM HISDB.dbo.MBMGRPSubsidiary WHERE MBMGRPCompany = @key;
    ELSE IF @kind = 'departments'
        SELECT RTRIM(o.MBMDepartment) FROM HISDB.dbo.MBM m JOIN HISDB.dbo.MBMOthers o ON m.MBMNumber = o.MBMNumber
        WHERE m.HSCLNInd <> 'H' AND o.MBMGroupCompany = @key GROUP BY o.MBMDepartment;
END
GO
