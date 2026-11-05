-- LG & Audit > Audit review (frmAuditInvoice): one case and its audit details. The update reuses cis.InvAdj_Audit, Prev / Next reuse
-- cis.InvEnq_Step.

-- CmdGo_Click: SP_CISAuditdetails (already parameterised) and the diagnosis descriptions (DiagDesc).
CREATE OR ALTER PROCEDURE cis.Audit_Case @casNumber varchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.SP_CISAuditdetails @CASNumber = @casNumber;
END
GO
