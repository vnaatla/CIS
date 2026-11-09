-- Claim Bordx > Print Bordx > BSN-PrintSummary (frmBSNPrintSummary): CISWorksheet SP_Bordxdetails case 1 (BANK SIMPANAN NASIONAL bordereaux) for one or more
-- bordx numbers (@bordxNos comma-separated).
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxBsn_Summary @bordxNos nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@bordxNos, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
    DECLARE @b TABLE (BordxRefNo nvarchar(20) PRIMARY KEY);
    INSERT @b SELECT DISTINCT t.c.value('.', 'nvarchar(20)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(20)') <> '';
    SELECT Bordx.BordxRefNo, COUNT(*) AS NOOFCLAIMS, SUM(Claim.CLMBordxAmt) AS CLMBordxAmt, Bordx.BordxDate
    FROM CISWorksheet.dbo.Bordx
    INNER JOIN CISWorksheet.dbo.Claim ON Claim.BordxRefNo = Bordx.BordxRefNo
    INNER JOIN dbo.CISCas ON CISCas.CASNumber = Claim.CASNumber
    INNER JOIN cis.Hdb_MBMOthers MBMOthers ON MBMOthers.MBMNumber = CISCas.MBMnumber
    WHERE MBMOthers.MBMGroupCompany = 'BANK SIMPANAN NASIONAL' AND Bordx.BordxRefNo IN (SELECT BordxRefNo FROM @b)
    GROUP BY Bordx.BordxRefNo, Bordx.BordxDate
    ORDER BY Bordx.BordxRefNo ASC, Bordx.BordxDate ASC;
END
GO
