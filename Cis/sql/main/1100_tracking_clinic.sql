-- Tracking > Clinic (FrmTracking): HISDB MBMTracking case 182 (SELECT * FROM CLN) with the pasted criteria replaced by parameters, and the combos.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.TrkCln_Search @code nvarchar(10) = NULL, @name nvarchar(200) = NULL, @group nvarchar(10) = NULL, @postcode nvarchar(10) = NULL,
    @city nvarchar(100) = NULL, @state nvarchar(100) = NULL, @contact nvarchar(200) = NULL, @doctor nvarchar(200) = NULL, @initial nvarchar(100) = NULL,
    @bnfType nvarchar(1) = NULL, @receivedFrom datetime = NULL, @receivedTo datetime = NULL, @panel nvarchar(1) = NULL, @type nvarchar(1) = NULL,
    @xray nvarchar(1) = NULL, @forRef nvarchar(50) = NULL, @location nvarchar(200) = NULL, @ffs nvarchar(1) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM cis.Hdb_CLN
    WHERE (@code IS NULL OR CLNCode = @code) AND (@name IS NULL OR CLNName LIKE '%' + @name + '%') AND (@group IS NULL OR ClinicGrpCode LIKE @group + '%')
      AND (@postcode IS NULL OR CLNPostCode = @postcode) AND (@city IS NULL OR CLNCity = @city) AND (@state IS NULL OR CLNState = @state)
      AND (@contact IS NULL OR CLNRepresentative LIKE '%' + @contact + '%') AND (@doctor IS NULL OR CLNDocName LIKE '%' + @doctor + '%')
      AND (@initial IS NULL OR CLNInitial LIKE '%' + @initial + '%') AND (@bnfType IS NULL OR CLNBNFType = @bnfType)
      AND (@receivedFrom IS NULL OR CLNReceivedDate >= @receivedFrom) AND (@receivedTo IS NULL OR CLNReceivedDate <= @receivedTo)
      AND (@panel IS NULL OR CLNPanel = @panel) AND (@type IS NULL OR CLNType = @type) AND (@xray IS NULL OR CLNXRay = @xray)
      AND (@forRef IS NULL OR CLNForRef = @forRef)
      AND (@location IS NULL OR CLNAdd1 LIKE '%' + @location + '%' OR CLNAdd2 LIKE '%' + @location + '%' OR CLNAdd3 LIKE '%' + @location + '%')
      AND (@ffs IS NULL OR CLNFFSStatus = @ffs)
    OPTION (RECOMPILE);
END
GO

-- CLNComboListing: clinic groups, cities, states.
CREATE OR ALTER PROCEDURE cis.TrkCln_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(ClinicGrpCode) + ' - ' + RTRIM(ClinicGrpName) FROM HISDB.dbo.ClinicGrp;
    SELECT RTRIM(CTYDescription) FROM HISMaintenance.dbo.CTY;
    SELECT RTRIM(STADescription) FROM HISDB.dbo.STA;
END
GO
