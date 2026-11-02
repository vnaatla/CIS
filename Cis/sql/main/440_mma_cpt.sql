-- Invoice > Other > MMA/CPT (frmMMASchedule) and its detail pop-up (frmShowInvestigation): HISMaintenance VIEW_MMA_Schedule.

-- CmdSearch_Click: every filter optional. The category filters use the text without its first character, as the desktop (Mid(x, 2, 8)).
CREATE OR ALTER PROCEDURE cis.Mma_Search
    @mainBp nvarchar(50) = NULL, @opcs nvarchar(50) = NULL, @procedure nvarchar(200) = NULL, @surgCat nvarchar(20) = NULL, @anaesCat nvarchar(20) = NULL,
    @part nvarchar(10) = NULL, @pageFrom nvarchar(20) = NULL, @pageTo nvarchar(20) = NULL, @cptCode nvarchar(50) = NULL, @cptDesc nvarchar(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT ProcMainBPCode, OPCSCode, OPCSNarrative, PartACost, CPTCode, CPTDesc, ProcSurgCat, ProcAnaesCat, ProcPageNo, ProcPart, SurgeonsFees,
           AnaesthetistsFees, Surgeon13thNew, Anesthetist13thNew
    FROM cis.Maint_VIEW_MMA_Schedule
    WHERE (@mainBp IS NULL OR ProcMainBPCode LIKE '%' + @mainBp + '%') AND (@opcs IS NULL OR OPCSCode LIKE '%' + @opcs + '%')
      AND (@procedure IS NULL OR OPCSNarrative LIKE '%' + @procedure + '%') AND (@surgCat IS NULL OR ProcSurgCat LIKE '%' + @surgCat + '%')
      AND (@anaesCat IS NULL OR ProcAnaesCat LIKE '%' + @anaesCat + '%') AND (@part IS NULL OR ProcPart = @part)
      AND (@pageFrom IS NULL OR ProcPageNo >= @pageFrom) AND (@pageTo IS NULL OR ProcPageNo <= @pageTo)
      AND (@cptCode IS NULL OR CPTCode = @cptCode) AND (@cptDesc IS NULL OR CPTDesc LIKE '%' + @cptDesc + '%');
END
GO
