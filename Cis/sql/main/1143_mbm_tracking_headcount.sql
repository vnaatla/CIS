-- MBM Tracking (Simple) variants: Head Count.
SET QUOTED_IDENTIFIER ON;
GO

-- MBM Tracking (Simple) > Head Count (bindHeadcount): the active principals / dependants per payor and policy-expiry month (HISDB ActiveHeadcount_CIS).
CREATE OR ALTER PROCEDURE cis.MbmTrk_Headcount @from date, @to date AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @f nvarchar(15) = CONVERT(nvarchar(10), @from, 23), @t nvarchar(15) = CONVERT(nvarchar(10), @to, 23);
    EXEC HISDB.dbo.ActiveHeadcount_CIS @FromDt = @f, @ToDt = @t;
END
GO
