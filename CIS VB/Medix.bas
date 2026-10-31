Attribute VB_Name = "Medix"
Option Explicit
Global CompanyName As String
Global CompanyAddress As String
Global BusinessYear As Integer
Global Access As String
Global Logon As String

Public Sub ActionLogUpdate(actionId As Integer, moduleName As String) 'Method added by SRI
    Dim rstActionLog As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "Insert into hisdb..tblactionlog values ('" + Logon + "'," + CStr(actionId) + ",''" + ",''" + ",''" + ",getdate()" + ",'CIS'" + ",'" + moduleName + "')"
    rstActionLog.Open strSql, medixmasterconnA, adOpenKeyset, adLockOptimistic
End Sub


Public Sub ErrorLog(FrmName As String, MethdName As String, ErrDesc As String)
Dim AL As New ADODB.Recordset
Dim strSql As String

    '1. Log in Error Details
        ' Error Details
    
    medixmasterconnA.Close
    medixmasterconn.Close
    medixmasterconnWS.Close
    medixmasterconnPayment.Close
    MediConnCISWS.Close
    MediConnCISPayment.Close
    MediConnCISDB.Close
    
    strSql = ""
    strSql = "Insert into Tbl_ErrorLog (SysName,FormName,MethodName,ErrDesc,ErrTime) values ('CIS','" + FrmName + "','" + MethdName + "','" + ErrDesc + "',GETDATE() )"
    AL.Open strSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    
    strSql = ""
    strSql = "Update USR set CIScnt=CIScnt-1 Where USRCode ='" & Trim(Logon) & "' And USRStatus ='A'"
    AL.Open strSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
    
    medixmasterconnMaint.Close
    
    Set medixmasterconnA = Nothing
    Set medixmasterconn = Nothing
    Set medixmasterconnWS = Nothing
    Set medixmasterconnPayment = Nothing
    Set medixmasterconnMaint = Nothing
    Set MediConnCISWS = Nothing
    Set MediConnCISPayment = Nothing
    Set MediConnCISDB = Nothing
    
    
    
    'Email to IT about Error Description
    Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
    Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
    Const cdoAnonymous = 0 'Do not authenticate
    Const cdoBasic = 1 'basic (clear-text) authentication
    Const cdoNTLM = 2 'NTLM
    Dim objMessage As Object
 '   On Error GoTo debugs
    Set objMessage = CreateObject("CDO.Message")
    objMessage.Subject = "Error Occured System: CIS, Form Name: " + FrmName + ", Method Name: " + MethdName
    objMessage.From = "notifications@medix.com.my"
    objMessage.To = "support.it@medix.com.my"
    objMessage.CC = "Bryan@medix.com.my"
    objMessage.TextBody = "Dear IT," & vbCrLf & vbCrLf & _
                          "Exception Raised in Above Module Please assist To check this Issue." & vbCrLf & vbCrLf & _
                          "Exception Description:" + ErrDesc & vbCrLf & vbCrLf & _
                          "User Name:" + Logon & vbCrLf & vbCrLf & _
                          "Thank you." & vbCrLf & _
                          "Regards, " & vbCrLf & _
                          "MediExpress (Malaysia) Sdn Bhd" & vbCrLf & vbCrLf
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.mediexpress.com.my"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "autoemail@mediexpress.com.my"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Medix@123"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    objMessage.Configuration.Fields.update
    objMessage.Send
    
'debugs:
'    If Err.Description <> "" Then
'        MsgBox Err.Description
'    End If
    
    
    
MsgBox ("System encounter Some Issue. Please Login Again")
Application.Quit

End

End Sub




Public Function GetCLNAnnualLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
Dim AL As New ADODB.Recordset
Dim strSql As String
          
    strSql = "Select AnnLimit from AnnualLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
             " AND AnnEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnEffDate  desc"
    AL.Open strSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If AL.RecordCount Then
         GetCLNAnnualLimit = AL!AnnLimit
    Else
         GetCLNAnnualLimit = 0
'         MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    AL.Close
    Set AL = Nothing
End Function


Public Function GenerateChkDigits()
Dim strSql As String
Dim SEQDigits As New ADODB.Recordset
Dim tmpSEQDigits As String
    strSql = "SELECT MBMChkDigits From MBMCheckDigits "
    SEQDigits.Open strSql, medixmasterconn, adOpenStatic, adLockOptimistic
        If SEQDigits!MBMChkDigits <> "99" Then
            tmpSEQDigits = Format(SEQDigits!MBMChkDigits + 1, "00")
            SEQDigits!MBMChkDigits = Format(SEQDigits!MBMChkDigits + 1, "00")
            SEQDigits.update
        Else
            tmpSEQDigits = "01"
            SEQDigits!MBMChkDigits = tmpSEQDigits
            SEQDigits.update
        End If
    SEQDigits.Close
    Set SEQDigits = Nothing
        GenerateChkDigits = tmpSEQDigits
End Function
Public Function FormLoadedByName(FormName As String) As Boolean
Dim i As Integer, fnamelc As String
    fnamelc = LCase$(FormName)
    FormLoadedByName = False
    For i = 0 To Forms.Count - 1
        If LCase$(Forms(i).Name) = fnamelc Then
            FormLoadedByName = True
            Exit Function
        End If
    Next
End Function

Public Function AgeCodeAutoNO(AgeCode)
Dim AGE As New ADODB.Recordset
Dim AGEInsert As New ADODB.Recordset
Dim strSql As String
Dim AGENumber As String

AGENumber = "01"
strSql = "select * from AGE where AgeCode = '" & Trim(AgeCode) & "'"
AGE.Open strSql, medixmasterconn, adOpenStatic, adLockBatchOptimistic
'currentNo = 1

If AGE.RecordCount = 0 Then
AGEInsert.Open "Select * from AGE", medixmasterconn, adOpenStatic, adLockOptimistic
With AGEInsert
    .AddNew
    !AgeCode = AgeCode
    .update
    End With
    AGEInsert.Close
    Set AGEInsert = Nothing
    End If
    AgeCodeAutoNO = AgeCode
    AGE.Close
Set AGE = Nothing

End Function


Public Function CompanyInfo()
    Dim COM As New ADODB.Recordset

    COM.Open "Medi", medixmasterconn, adOpenKeyset, adLockOptimistic

    CompanyName = COM!COMName
    CompanyAddress = COM!COMAddress
    BusinessYear = COM!COMCurrentYear

    COM.Close
End Function

Public Function GetAge(EffectiveDate As Date, DateOfBirth As Date)
    Dim CheckDate As Long
    Dim CheckYear As Integer
    Dim CheckDOB As Integer
    
    CheckDate = Year(EffectiveDate) - Year(DateOfBirth)
        If CheckDate <= 17 Then
            If CheckDate <= 0 Then
                If CheckDate <= 0 Then
                    GetAge = "-1D"
                Else
                    GetAge = DateDiff("d", EffectiveDate, DateOfBirth) & "D"
                End If
            Else
                If Month(DateOfBirth) > Month(EffectiveDate) Then
                    GetAge = CheckYear - 1 & "Y"
                Else
                    If Day(DateOfBirth) > Day(EffectiveDate) Then
                        GetAge = CheckYear - 1 & "Y"
                    Else
                        GetAge = CheckYear & "Y"
                    End If
                    
                End If
            End If
    Else
        CheckYear = CheckDate
        If Month(DateOfBirth) > Month(EffectiveDate) Then
            GetAge = CheckYear - 1 & "Y"
        Else
            If Day(DateOfBirth) > Day(EffectiveDate) Then
                GetAge = CheckYear - 1 & "Y"
            Else
                GetAge = CheckYear & "Y"
            End If
            
        End If
    End If
End Function

Public Function GetAgeBand(GetAgeValue As String)
    On Error Resume Next
    Dim strSql As String
    Dim AgeOnly As Integer
    Dim AGE As New ADODB.Recordset
    Dim AgeCat As String
    Dim AgeLenght As String
    
    AgeLenght = Len(GetAgeValue)
        
    If Mid(GetAgeValue, AgeLenght, 1) = "Y" Then
        AgeCat = "Y"
        AgeOnly = Mid(GetAgeValue, 1, InStr(1, GetAgeValue, "Y") - 1)
        strSql = "Select AGECODE ,AGEDescription from AGE Where cast(substring(AGEfrom ,1  , LEN(AGEFROM)-1   ) as integer) <='" & AgeOnly & "' and  cast(substring(AGETo ,1  , LEN(AGETo)-1   ) as integer) >='" & AgeOnly & "'"
    Else
        AgeCat = "D"
        AgeOnly = Mid(GetAgeValue, 1, InStr(1, GetAgeValue, "D") - 1)
        strSql = "Select AGECODE ,AGEDescription,  AgeFrom from AGE Where cast(substring(AGEfrom ,1  , LEN(AGEFROM)-1   ) as integer) <= '" & AgeOnly & "' order by AgeFrom"
    End If
    AGE.MaxRecords = 1
    AGE.Open strSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AGE.RecordCount Then
        If AgeCat = "D" Then
             AgeLenght = Len(AGE!AGEFrom)
              If AgeOnly >= Mid(AGE!AGEFrom, 1, AgeLenght - 1) Then
                GetAgeBand = AGE!AgeCode & " - " & AGE!AGEDescription
            End If
        Else
            GetAgeBand = AGE!AgeCode & " - " & AGE!AGEDescription
        End If
    Else
        MsgBox "Age Band not found or Date out of Acceptable Range! " & _
            vbNewLine & "Please verify your maintenance table!", vbCritical + vbOKOnly
    End If
    AGE.Close
    Set AGE = Nothing
End Function
Public Function PopulatePendingReasons(strtmpPayto As String)
    Dim varpipe() As String
    Dim Record As ListItem
    frmCasePending.lstpendinglist.ListItems.Clear
    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim rstPendingReasons As New ADODB.Recordset
    
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "spGetPendingReasons_OP"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("Category", adVarChar, adParamInput, 2000, strtmpPayto)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("Active", adInteger, adParamInput, 2000, 1)
    cmd.Parameters.Append prm2
    rstPendingReasons.CursorLocation = adUseClient
    rstPendingReasons.CursorType = adOpenForwardOnly
    rstPendingReasons.CacheSize = 100
    Set rstPendingReasons = cmd.Execute
    
    varpipe = Split(strPendingReasonId, ",")
    If (Len(strPendingReasonId) > 0) Then
    
    Else
    Do While Not rstPendingReasons.EOF
    
                Set Record = frmCasePending.lstpendinglist.ListItems.Add(, , "")
                Record.SubItems(1) = IIf(Len(Trim(rstPendingReasons!PendReasonDesc)), Trim(rstPendingReasons!PendReasonDesc), "")
                Record.SubItems(2) = IIf(Len(Trim(rstPendingReasons!PendReasonCode)), Trim(rstPendingReasons!PendReasonCode), "")
            rstPendingReasons.MoveNext
            Loop
    End If
    
    rstPendingReasons.Close
    Set rstPendingReasons = Nothing
End Function

Public Function GetPendingReasonsID(caseNo As String, latest As Integer) As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rstPendingReasons As New ADODB.Recordset

Set cmd.ActiveConnection = MediConnCISDB
cmd.CommandText = "spGetPendingReasonsByCase"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 300
Set Prm1 = cmd.CreateParameter("casNo", adVarChar, adParamInput, 2000, caseNo)
cmd.Parameters.Append Prm1
Set prm2 = cmd.CreateParameter("latest", adInteger, adParamInput, 2000, latest)
cmd.Parameters.Append prm2
rstPendingReasons.CursorLocation = adUseClient
rstPendingReasons.CursorType = adOpenForwardOnly
rstPendingReasons.CacheSize = 100
Set rstPendingReasons = cmd.Execute

Do While Not rstPendingReasons.EOF
            
            GetPendingReasonsID = IIf(Len(Trim(rstPendingReasons!ReasonID)), Trim(rstPendingReasons!ReasonID), "")
            
        rstPendingReasons.MoveNext
        Loop


rstPendingReasons.Close
Set rstPendingReasons = Nothing

End Function

Public Function GetPremium(MBMType As String, InsureType As String, PayorValue As String, AgeValue As String, PlanValue As String, HealthValue As String, PayorEffDate As Date) As Double
    Dim prm As New ADODB.Recordset
    Dim strSql As String
    Dim PLnRec As New ADODB.Recordset
    Dim PlnSql As String
    Dim NoPremium As Boolean

    NoPremium = True
    PlnSql = "SELECT PRMInd from PLN WHERE PLNCode='" & Trim(PlanValue) & _
             "' AND PAYCode='" & Trim(PayorValue) & "'"
    PLnRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not PLnRec.EOF
        If Trim(PLnRec!PRMInd) <> "" Then
            If Trim(PLnRec!PRMInd) = "Y" Then
                NoPremium = False
            End If
        End If
        PLnRec.MoveNext
    Loop
    PLnRec.Close
    Set PLnRec = Nothing
    GetPremium = 0
    If NoPremium = False Then

        strSql = "SELECT PRMAmount, SuppPRMAmt From PRM "
        strSql = strSql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') "
        strSql = strSql & "AND (PLNCode = '" & Trim(PlanValue) & "') AND (AGECode = '" & Trim(AgeValue) & "') "
        strSql = strSql & "AND (PRMEffDate  <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "') "
        If Trim(HealthValue) <> "S" Then
            strSql = strSql & "AND (PAYCode = '" & PayorValue & "')"
        End If
        prm.Open strSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        If prm.RecordCount Then
            If Trim(MBMType) = "P" Then
                GetPremium = IIf(IsNumeric(prm!PRMAmount), prm!PRMAmount, 0)
            ElseIf Trim(MBMType) = "S" Then
                GetPremium = IIf(IsNumeric(prm!SuppPrmAmt), prm!SuppPrmAmt, 0)
            End If
        Else
            MsgBox "Age Out of Premium Range! Please check your Premium Table."
        End If
            
        prm.Close
        Set prm = Nothing
    End If
End Function

Public Function GetMCO(tmpMBMTYpe As String, InsureType As String, HealthValue As String, AdultOrChild As String, PayorEffDate As Date, PLNCode As String, tmpPayor As String) As Double
    Dim MCO As New ADODB.Recordset
    Dim tmpSql, strSql As String
    Dim MCONew As New ADODB.Recordset
    Dim strsqlNEW As String
    Dim PLnRec As New ADODB.Recordset
    Dim PlnSql As String
    Dim NoMCO As Boolean

    NoMCO = True
    PlnSql = "SELECT MCOInd from PLN WHERE PLNCode='" & Trim(PLNCode) & _
             "' AND PAYCode='" & Trim(tmpPayor) & "'"
    PLnRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not PLnRec.EOF
        If Trim(PLnRec!MCOInd) <> "" Then
            If Trim(PLnRec!MCOInd) = "Y" Then
                NoMCO = False
            End If
        End If
        PLnRec.MoveNext
    Loop
    PLnRec.Close
    Set PLnRec = Nothing
    GetMCO = 0
    
    If NoMCO = False Then
        If Trim(AdultOrChild) = "A" Then
            If Trim(tmpMBMTYpe) = "P" Then
                tmpSql = "SELECT MCOAmtAdult As MCOAmt From MCO "
            ElseIf Trim(tmpMBMTYpe) = "S" Then
                tmpSql = "SELECT MCOSuppAdultAmt As MCOAmt From MCO "
            End If
        Else
            If Trim(tmpMBMTYpe) = "P" Then
                tmpSql = "SELECT MCOAmtChild As MCOAmt From MCO "
            ElseIf Trim(tmpMBMTYpe) = "S" Then
                tmpSql = "SELECT MCOSuppChildAmt As MCOAmt From MCO "
            End If
        End If
    
        strsqlNEW = tmpSql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') AND (PLNCode = '" & Trim(PLNCode) & "') " & _
                " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                    
        MCONew.Open strsqlNEW, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        strSql = tmpSql & "WHERE (INSCode = '" & Trim(InsureType) & "') AND (HLTCode = '" & Trim(HealthValue) & "') " & _
                " AND MCOEffdate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                    
        MCO.Open strSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If MCONew.RecordCount Then
            GetMCO = IIf(IsNumeric(MCONew!MCOAmt), MCONew!MCOAmt, 0)
        ElseIf MCO.RecordCount Then
            GetMCO = IIf(IsNumeric(MCO!MCOAmt), MCO!MCOAmt, 0)
        Else
            MsgBox "MCO Information Is Not Correct" & vbNewLine & " Please Check your MCO table", vbOKOnly + vbCritical, DialogBoxTitle
        End If
        MCONew.Close
        Set MCONew = Nothing
        MCO.Close
        Set MCO = Nothing
    End If
End Function

'==================== Get CIS Principal  & Supplementary Limit ===========================================
Public Function GetAnnualLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, MBMType As String, tmpPayor, tmpLimitStatus) As String

Dim AL As New ADODB.Recordset
Dim strSql As String
Dim PLnRec As New ADODB.Recordset
Dim PlnSql As String
Dim Unlimited As Boolean
Unlimited = False
Dim isPrinDepGPShared As Boolean

    Unlimited = False
    isPrinDepGPShared = False
    PlnSql = ""
    PlnSql = "SELECT AnnLmtIND as AnnLimitIND, ANNLmtIndSupp, COALESCE(isPrinDepGPShared,0) AS isPrinDepGPShared,ChildLimitInd from PLN WHERE PLNCode='" & Trim(PlanValue) & _
             "' AND PAYCode='" & Trim(tmpPayor) & "'"
    'If Trim(MBMType) <> "P" Then
    '    PlnSql = "SELECT ANNLmtIndSupp as AnnLimitIND from PLN WHERE PLNCode='" & Trim(PlanValue) & _
    '         "' AND PAYCode='" & Trim(tmpPayor) & "'"
    'End If
    PLnRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not PLnRec.EOF
        isPrinDepGPShared = PLnRec!isPrinDepGPShared
        If MBMType = "P" Then
            If Trim(PLnRec!AnnLimitIND) <> "" Then
                If Trim(PLnRec!AnnLimitIND) = "N" Then
                    Unlimited = True
                ElseIf Trim(PLnRec!AnnLimitIND) = "Y" Then
                    Unlimited = False
                Else
                    Unlimited = False
                    'MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
                End If
            Else
                Unlimited = False
                'MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
            End If
        Else
            If tmpLimitStatus <> "B" Then
                If Trim(PLnRec!ANNLmtIndSupp) <> "" Then
                    If Trim(PLnRec!ANNLmtIndSupp) = "N" Then
                        Unlimited = True
                    ElseIf Trim(PLnRec!ANNLmtIndSupp) = "Y" Then
                        Unlimited = False
                    Else
                        Unlimited = False
                        'MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
                    End If
                Else
                    Unlimited = False
                    'MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
                End If
            Else
                Unlimited = False
            End If
        End If
        PLnRec.MoveNext
    Loop
    PLnRec.Close
    Set PLnRec = Nothing
    GetAnnualLimit = "FULL"
    
    If Unlimited = False Then
        strSql = "Select AnnLimit, SuppLimit, SuppLimitStatus,ChildDepLimit,coalesce(AnnualLimit.asolimit,0) As asolimit from AnnualLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
                 " AND AnnEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by AnnEffDate  desc"
        AL.Open strSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        If Not AL.EOF Then
            tmpLimitStatus = AL!SuppLimitStatus
            If (IIf(Len(AL!SuppLimitStatus), AL!SuppLimitStatus, "")) = "" Then
                GetAnnualLimit = 0
                MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
            ElseIf (IIf(Len(AL!SuppLimitStatus), AL!SuppLimitStatus, "")) = "D" Then
                GetAnnualLimit = "FULL"
            Else
                GetAnnualLimit = (IIf(IsNumeric(AL!AnnLimit), AL!AnnLimit, 0) + IIf(IsNumeric(AL!ASOlimit), AL!ASOlimit, 0))
            End If
            
            If Trim(MBMType) = "S" Then
                If isPrinDepGPShared Then
                    GetAnnualLimit = (IIf(IsNumeric(AL!AnnLimit), AL!AnnLimit, 0) + IIf(IsNumeric(AL!ASOlimit), AL!ASOlimit, 0))
                Else
                    If Trim(AL!SuppLimitStatus) <> "A" Then
                        GetAnnualLimit = IIf(IsNumeric(AL!SuppLimit), AL!SuppLimit, 0)
                    End If
                End If
            End If
        Else
             GetAnnualLimit = 0
            ' MsgBox "Annual Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
        End If
        AL.Close
        Set AL = Nothing
    End If
End Function

Public Function GetSpecLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, MBMType As String, tmpPayor) As String

Dim VL As New ADODB.Recordset
Dim strSql As String
Dim PLnRec As New ADODB.Recordset
Dim PlnSql As String

Dim SLmtInd As String
SLmtInd = "N"


    PlnSql = "SELECT PSpecLimit AS SpecLmtIND from PLN WHERE PLNCode='" & Trim(PlanValue) & _
             "' AND PAYCode='" & Trim(tmpPayor) & "'"
    PLnRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not PLnRec.EOF
        If Trim(PLnRec!SpecLmtIND) <> "" Then
            If Trim(PLnRec!SpecLmtIND) = "F" Then
                SLmtInd = "F"
            ElseIf Trim(PLnRec!SpecLmtIND) = "N" Then
                SLmtInd = "N"
            ElseIf Trim(PLnRec!SpecLmtIND) = "Y" Then
                SLmtInd = "Y"
            Else
                SLmtInd = "N"
            End If
        Else
            SLmtInd = "N"
        End If
        PLnRec.MoveNext
    Loop
    PLnRec.Close
    Set PLnRec = Nothing
    
    If SLmtInd = "F" Then
        GetSpecLimit = "FULL"
    ElseIf SLmtInd = "N" Then
        GetSpecLimit = "N/A"
    ElseIf SLmtInd = "Y" Then
        strSql = ""
        strSql = "Select PrinLimit,SuppLimit, SuppLimitStatus from SpecialistLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "'"
        VL.Open strSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
        If Not VL.EOF Then
            If (IIf(Len(VL!SuppLimitStatus), VL!SuppLimitStatus, "")) = "" Then
                GetSpecLimit = "FULL"
                MsgBox "Specialist Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
            ElseIf (IIf(Len(VL!SuppLimitStatus), VL!SuppLimitStatus, "")) = "D" Then
                GetSpecLimit = "FULL"
            Else
                GetSpecLimit = IIf(IsNumeric(VL!PrinLimit), VL!PrinLimit, 0)
            End If
            
          
            If Trim(MBMType) = "S" Then
                If Trim(VL!SuppLimitStatus) <> "A" Then
                    GetSpecLimit = IIf(IsNumeric(VL!SuppLimit), VL!SuppLimit, 0)
                End If
            End If
        Else
             GetSpecLimit = 0
            ' MsgBox "Specialist Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
        End If
        VL.Close
        Set VL = Nothing
    End If
    
End Function

'Public Function GetSpecLimit2(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, MBMType As String, tmpPayor) As String

'Dim VL As New ADODB.Recordset
'Dim strsql As String
'Dim PLnRec As New ADODB.Recordset
'Dim PlnSql As String
'Dim Unlimited As Boolean

'    Unlimited = False
'    PlnSql = "SELECT PSpecLimit AS VisitLmtIND from PLN WHERE PLNCode='" & Trim(PlanValue) & _
             "' AND PAYCode='" & Trim(tmpPayor) & "'"
    'If Trim(MBMType) <> "P" Then
    '    PlnSql = "SELECT VISLmtIndSupp AS VisitLmtIND from PLN WHERE PLNCode='" & Trim(PlanValue) & _
    '         "' AND PAYCode='" & Trim(tmpPayor) & "'"
    'End If
'    PLnRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
'    Do While Not PLnRec.EOF
'        If Trim(PLnRec!VisitLmtIND) <> "" Then
'            If Trim(PLnRec!VisitLmtIND) = "N" Then
'                Unlimited = True
'            End If
'        End If
'        PLnRec.MoveNext
'    Loop
'    PLnRec.Close
'    Set PLnRec = Nothing
'    GetSpecLimit = "FULL"
    
'    If Unlimited = False Then
'        strsql = ""
'        strsql = "Select PrinLimit, VisitLimitStatus from SpecialistLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "'"
'        VL.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
'        If Not VL.EOF Then
'            GetSpecLimit = IIf(IsNumeric(VL!PrinLimit), VL!PrinLimit, 0)
            'If Trim(MBMType) = "S" Then
            '    If Trim(VL!SuppVisitLimitStatus) <> "A" Then
            '        GetVisitLimit = IIf(IsNumeric(VL!SuppVisitLimit), VL!SuppVisitLimit, 0)
            '    End If
            'End If
'        Else
'             GetSpecLimit = 0
'             MsgBox "Specialist Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
'        End If
'        VL.Close
'        Set VL = Nothing
'    End If
'End Function
'==================== Get CIS Principal & Supplementary Limit ===========================================
Public Function GetVisitLimit(InsureType As String, PlanValue As String, HealthValue As String, PayorEffDate As Date, MBMType As String, tmpPayor) As String

Dim VL As New ADODB.Recordset
Dim strSql As String
Dim PLnRec As New ADODB.Recordset
Dim PlnSql As String
Dim Unlimited As Boolean

    Unlimited = False
    PlnSql = "SELECT VISLmtInd AS VisitLmtIND,ChildLimitInd from PLN WHERE PLNCode='" & Trim(PlanValue) & _
             "' AND PAYCode='" & Trim(tmpPayor) & "'"
    'If Trim(MBMType) <> "P" Then
    '    PlnSql = "SELECT VISLmtIndSupp AS VisitLmtIND from PLN WHERE PLNCode='" & Trim(PlanValue) & _
    '         "' AND PAYCode='" & Trim(tmpPayor) & "'"
    'End If
    PLnRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not PLnRec.EOF
        If Trim(PLnRec!VisitLmtIND) <> "" Then
            If Trim(PLnRec!VisitLmtIND) = "N" Then
                Unlimited = True
            ElseIf Trim(PLnRec!VisitLmtIND) = "Y" Then
                 Unlimited = False
            Else
                 Unlimited = False
            End If
        Else
            Unlimited = False
        End If
        PLnRec.MoveNext
    Loop
    PLnRec.Close
    Set PLnRec = Nothing
    GetVisitLimit = "FULL"
    
    If Unlimited = False Then
        
        strSql = "Select VisitLimit, SuppVisitLimit, SuppVisitLimitStatus from VisitLimit Where INSCode ='" & Trim(InsureType) & "' And PLNCode ='" & Trim(PlanValue) & "' And HLTCode = '" & Trim(HealthValue) & "' " & _
                 " AND VisitEffDate <= '" & Format(PayorEffDate, "mm/dd/yyyy") & "' order by VisitEffDate  desc"
        VL.Open strSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
        If Not VL.EOF Then
            GetVisitLimit = IIf(IsNumeric(VL!VisitLimit), VL!VisitLimit, 0)
            If Trim(MBMType) = "S" Then
                If Trim(VL!SuppVisitLimitStatus) <> "A" Then
                    GetVisitLimit = IIf(IsNumeric(VL!SuppVisitLimit), VL!SuppVisitLimit, 0)
                End If
            End If
        Else
             GetVisitLimit = 0
             'MsgBox "Visit Limit Information Is Not Correct! " & vbNewLine & "Please check you Maintenance Table", vbOKOnly + vbCritical, DialogBoxTitle
        End If
        VL.Close
        Set VL = Nothing
    End If
End Function
'==================== Get CIS Principal & Supplementary Limit ===========================================


Public Function GetProviderCharges(ServiceType As String, InsureType As String, HealthValue As String, AdultChild As String) As Double
On Error Resume Next
    Dim PRV As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select * from PRV Where SRVCode ='" & ServiceType & "' And INSCode ='" & InsureType & "' And HLTCode ='" & HealthValue & "'"
    PRV.Open strSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If PRV.RecordCount = 0 Then
        MsgBox "Provider Charges Information Not Correct", vbCritical + vbOKOnly, DialogBoxTitle
    Else
            GetProviderCharges = PRV!PRVProviderCharges
    End If
    PRV.Close
    Set PRV = Nothing
End Function

Public Function CheckBatch(Payor As String, BordxDate As Date, BordxType As String) As String
    Dim BAT As New ADODB.Recordset
    Dim strSql As String
    Dim CurrentBatch As Integer
    
    strSql = "Select * from BAT Where PAYCode ='" & Payor & "' And BATBordxDate = '" & Format(BordxDate, "MM/DD/YYYY") & "' And BordxType = '" & BordxType & "'"
    BAT.Open strSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                    
    If BAT.RecordCount = 0 Then
        BAT.AddNew
        With BAT
        If BordxType = "N" Then
            !BordxType = "N"
        ElseIf BordxType = "E" Then
            !BordxType = "E"
        End If
            !PAYCode = Payor
            !BATBordxDate = BordxDate
            !BATSequence = "1"
        End With
        BAT.update
        CheckBatch = "1"
    Else
        With BAT
            CurrentBatch = !BATSequence + 1
            !BATSequence = CurrentBatch
            BAT.update
        End With
        CheckBatch = CurrentBatch
    End If
 
End Function

Public Function l(stringvalue As String, stringlenght As Integer)
Dim stringformat
Dim LengLoop As Integer

    For LengLoop = 1 To stringlenght
        stringformat = stringformat + "1"
    Next LengLoop
    If IsNumeric(stringvalue) Then
        RSet stringformat = stringvalue
    Else
        LSet stringformat = stringvalue
    End If
    l = stringformat
End Function


Public Function GenerateCaseNo(PayorCode)
    Dim CAQ As New ADODB.Recordset
    Dim strSql As String
    Dim CAQNumber As String
    Dim CAQCharacter As String
    
    strSql = "Select * from CAQ Where PAYCode = '" & PayorCode & "'"
    CAQ.Open strSql, medixmasterconn, adOpenStatic, adLockOptimistic
    
    If CAQ.RecordCount = 0 Then
        With CAQ
            .AddNew
            !PAYCode = PayorCode
            !CAQNumber = "0000001"
            .update
            CAQNumber = "0000001"
        End With
    Else
        CAQ!CAQNumber = Format(CAQ!CAQNumber + 1, "0000000")
        CAQNumber = CAQ!CAQNumber
        CAQ.update
    End If

    GenerateCaseNo = PayorCode & CAQNumber
    CAQ.Close
    Set CAQ = Nothing
End Function

Public Function GetAdjustmentNo(MembershipNo As String, Policyno As String, caseNo As String, AdjustmentType As String) As Integer
    Dim ADJ As New ADODB.Recordset
    Dim adjstr As String
    Dim NextSequence As Integer
    
    adjstr = "Select * from ADJ Where ADJMemberNo ='" & MembershipNo & "' "
    adjstr = adjstr & "And ADJPolicyNo ='" & Policyno & "' And ADJCaseNo ='" & caseNo & "' And ADJType ='" & AdjustmentType & "'"
    ADJ.Open adjstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If ADJ.RecordCount = 0 Then
        ADJ.AddNew
         With ADJ
            !ADJMemberno = MembershipNo
            !ADJPolicyNo = Policyno
            !ADJDate = Date
            !ADJType = AdjustmentType
            If AdjustmentType = "C" Then
                !ADJCaseNo = caseNo
            End If
            !ADJSequence = 1
        End With
        ADJ.update
        GetAdjustmentNo = 1
    Else
        NextSequence = ADJ!ADJSequence + 1
        ADJ!ADJSequence = NextSequence
        GetAdjustmentNo = NextSequence
        ADJ.update
    End If
            
End Function

Public Function GetLastCoveredID(MembershipNo As String) As Integer
On Error Resume Next
    Dim MBMCovered As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT MBMCNumber, MBMCMemberType, MAX(MBMCCoverID) AS CoveredID "
    strSql = strSql & "From MBMCoveredPersons GROUP BY MBMCNumber, MBMCMemberType "
    strSql = strSql & "HAVING (MBMCNumber = '" & MembershipNo & "') AND "
    strSql = strSql & "(MBMCMemberType = 'C')"
    MBMCovered.Open strSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    GetLastCoveredID = MBMCovered!CoveredID + 1
    
    MBMCovered.Close
    Set MBMCovered = Nothing
End Function

Public Function MCORefund(ExpiryDate As Date, CancelDate As Date, EffDate As Date, MCOFees As Double) As Double

    MCORefund = (CInt(ExpiryDate - CancelDate) / (ExpiryDate - EffDate)) * MCOFees
    If MCORefund < 0 Then
        MCORefund = 0
    End If
End Function

Public Function CheckAccess(AccessNo As Integer) As Boolean
On Error Resume Next
    If Mid(Access, AccessNo, 1) = 1 Then
        CheckAccess = True
    Else
        CheckAccess = False
    End If
End Function

Public Sub SearchForm(ScanCode As String)
On Error GoTo errOpen
    
    If CheckAccess(95) = True Then 'MCO Invoice
        If FormLoadedByName("frmMCOInvoiceNew") = True Then
            frmMCOInvoiceNew.SetFocus
        Else
            frmMCOInvoiceNew.Show
        End If
    
    ElseIf CheckAccess(96) = True Then  'MCO Fee Listing
        If FormLoadedByName("frmMCOFeeListNew") = True Then
            frmMCOFeeListNew.SetFocus
        Else
            frmMCOFeeListNew.Show
        End If
        
    ElseIf CheckAccess(98) = True Then 'Wks Reg.
        If FormLoadedByName("frmNewWSReg") = True Then
            frmNewWSReg.SetFocus
        Else
            frmNewWSReg.Show
        End If
        
    ElseIf CheckAccess(99) = True Then 'Wks Adjustment
       If FormLoadedByName("frmNewWSAjd") = True Then
            frmNewWSAjd.SetFocus
        Else
            frmNewWSAjd.Show
        End If
        
    ElseIf CheckAccess(100) = True Then 'Select Wks for Bordx
        If FormLoadedByName("frmWSBordx") = True Then
            frmWSBordx.SetFocus
        Else
            frmWSBordx.Show
        End If
        
    ElseIf CheckAccess(101) = True Then 'Remove Wks from Bordx
        If FormLoadedByName("FrmBordxRevise") = True Then
            FrmBordxRevise.SetFocus
        Else
            FrmBordxRevise.Show
        End If
    
    ElseIf CheckAccess(102) = True Then 'Print Bordx
        If FormLoadedByName("frmPrintBordx") = True Then
            frmPrintBordx.SetFocus
        Else
            frmPrintBordx.Show
        End If
    
    ElseIf CheckAccess(103) = True Then 'Bordx Send to Payor
        If FormLoadedByName("FrmSendtoPayor") = True Then
            FrmSendtoPayor.SetFocus
        Else
            FrmSendtoPayor.Show
        End If
        
    ElseIf CheckAccess(104) = True Then 'Bordx Return From Payor
        If FormLoadedByName("frmReturnPayor") = True Then
            FrmReturnPayor.SetFocus
        Else
            FrmReturnPayor.Show
        End If
        
    ElseIf CheckAccess(105) = True Then 'Bordx Send to MNRB
        If FormLoadedByName("FrmSendtoMNRB") = True Then
            FrmSendtoMNRB.SetFocus
        Else
            FrmSendtoMNRB.Show
        End If
        
    ElseIf CheckAccess(106) = True Then 'Bordx Send to A/C
        If FormLoadedByName("frmSendtoAC") = True Then
            FrmSendtoAC.SetFocus
        Else
            FrmSendtoAC.Show
        End If
        
    ElseIf CheckAccess(107) = True Then 'Borx A/C Period
        If FormLoadedByName("FrmACPeriod") = True Then
            FrmACPeriod.SetFocus
        Else
            FrmACPeriod.Show
        End If
        
    ElseIf CheckAccess(108) = True Then 'Bordx Control
        If FormLoadedByName("FrmBordxControl") = True Then
            FrmBordxControl.SetFocus
        Else
            FrmBordxControl.Show
        End If
        
    ElseIf CheckAccess(109) = True Then 'O/S Bordx
        If FormLoadedByName("FrmOSBordx") = True Then
            FrmOSBordx.SetFocus
        Else
            FrmOSBordx.Show
        End If
        
    ElseIf CheckAccess(110) = True Then 'Bordx & Wks
        If FormLoadedByName("FrmBordxWks") = True Then
            FrmBordxWks.SetFocus
        Else
            FrmBordxWks.Show
        End If
        
    ElseIf CheckAccess(111) = True Then 'Payment to Clinic
        If FormLoadedByName("frmPaymentClinic") = True Then
            frmPaymentClinic.SetFocus
        Else
            frmPaymentClinic.Show
        End If
    
    ElseIf CheckAccess(112) = True Then 'Receipt from Member
        If FormLoadedByName("FrmReceiptMember") = True Then
            FrmReceiptMember.SetFocus
        Else
            FrmReceiptMember.Show
        End If
        
    ElseIf CheckAccess(113) = True Then 'Receipt From MNRB
        If FormLoadedByName("FrmReceiptMNRB") = True Then
            FrmReceiptMNRB.SetFocus
        Else
            FrmReceiptMNRB.Show
        End If
        
    ElseIf CheckAccess(114) = True Then 'Receipt From payor
        If FormLoadedByName("FrmReceiptPayor") = True Then
            FrmReceiptPayor.SetFocus
        Else
            FrmReceiptPayor.Show
        End If
    
    ElseIf CheckAccess(115) = True Then 'Plan
        If FormLoadedByName("frmMaintenancePlan") = True Then
            frmMaintenancePlan.SetFocus
        Else
            frmMaintenancePlan.Show
        End If
        
    ElseIf CheckAccess(116) = True Then 'Case
        If FormLoadedByName("frmMaintenanceOthers") = True Then
            frmMaintenanceOthers.SetFocus
        Else
            frmMaintenanceOthers.Show
        End If
        
    ElseIf CheckAccess(117) = True Then 'Calculation
        If FormLoadedByName("frmMaintenanceCal") = True Then
            frmMaintenanceCal.SetFocus
        Else
            frmMaintenanceCal.Show
        End If
        
    ElseIf CheckAccess(118) = True Then 'Benefit
        If FormLoadedByName("frmMaintenanceBenefits") = True Then
            frmMaintenanceBenefits.SetFocus
        Else
            frmMaintenanceBenefits.Show
        End If
        
    ElseIf CheckAccess(119) = True Then 'Check Item
        If FormLoadedByName("frmCheckingItems") = True Then
            frmCheckingItems.SetFocus
        Else
            frmCheckingItems.Show
        End If
        
    Else
       MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
    End If

    Exit Sub
errOpen:
    If Err.Number = 7 Then
        MsgBox "Out Of Memory! Please close some of your forms! ", vbCritical
    Else
        MsgBox Err.Description
    End If


End Sub

Public Function MCOPremiumReduction(ExpiryDate As Date, CancelDate As Date, EffDate As Date, Premium As Double) As Double
    If DateDiff("d", ExpiryDate, CancelDate) < 0 Then
        MCOPremiumReduction = (CInt(ExpiryDate - CancelDate) / (ExpiryDate - EffDate)) * Premium
    Else
        MCOPremiumReduction = 0
    End If
End Function

Public Sub EnterToTab(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        If KeyAscii = 13 Then   'Turn off the beep
            KeyAscii = 0
        End If
        SendKeys "{tab}"
    End If
End Sub

Public Function InsertMCOFees(ExpiryDate As Date, EffectiveDate As Date, MCOFees As Double) As Double
    InsertMCOFees = (CDbl(ExpiryDate - EffectiveDate) * MCOFees) / 364
End Function

Public Function InsertPremium(ExpiryDate As Date, EffectiveDate As Date, Premium As Double) As Double
    InsertPremium = (CInt(ExpiryDate - EffectiveDate) / 364) * Premium
End Function

Public Function InsertProvider(ExpiryDate As Date, EffectiveDate As Date, AccessFees As Double) As Double
    InsertProvider = (CInt(ExpiryDate - EffectiveDate) / 364) * AccessFees
End Function

Public Function GenerateCLMBordxNo(PAYCode)
    Dim CLMBordxSEQ As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim strSql As String
    Dim strsql2 As String
    Dim CMLBordxNo As String
    Dim CAQCharacter As String
    Dim HLTCode As String
    
    strSql = "Select * from CLAIM_BordxSEQ Where PAYCode = '" & Trim(PAYCode) & "'"
    strsql2 = "Select * from PAY Where PAYCode = '" & Trim(PAYCode) & "'"
    CLMBordxSEQ.Open strSql, MediConnCISWS, adOpenStatic, adLockOptimistic
    PAY.Open strsql2, MediConnCISDB, adOpenStatic, adLockOptimistic
    
    If PAY.RecordCount > 0 Then
        HLTCode = PAY!HLTCode
    End If
    
    If CLMBordxSEQ.RecordCount = 0 Then
        With CLMBordxSEQ
            .AddNew
            !HLTCode = HLTCode
            !PAYCode = PAYCode
            !CLMBordxNo = "00001"
            .update
            CMLBordxNo = "00001"
        End With
    Else
        CLMBordxSEQ!CLMBordxNo = Format(CLMBordxSEQ!CLMBordxNo + 1, "00000")
        CMLBordxNo = CLMBordxSEQ!CLMBordxNo
        CLMBordxSEQ.update
    End If
    
    GenerateCLMBordxNo = "K" & PAYCode & CMLBordxNo
    CLMBordxSEQ.Close
    Set CLMBordxSEQ = Nothing
End Function


Public Function GenerateCLMBordxNoVOID(PAYCode)
    Dim CLMBordxSEQ As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim strSql As String
    Dim strsql2 As String
    Dim CMLBordxNo As String
    Dim CAQCharacter As String
    Dim HLTCode As String
    
    strSql = "Select * from CLAIM_BordxSEQ Where PAYCode = '" & Trim(PAYCode) & "'"
    CLMBordxSEQ.Open strSql, MediConnCISWS, adOpenStatic, adLockOptimistic
    
    If CLMBordxSEQ.RecordCount = 0 Then
        With CLMBordxSEQ
            .AddNew
            !HLTCode = HLTCode
            !PAYCode = PAYCode
            !CLMBordxNo = "00001"
            .update
            CMLBordxNo = "00001"
        End With
    Else
        CLMBordxSEQ!CLMBordxNo = Format(CLMBordxSEQ!CLMBordxNo + 1, "00000")
        CMLBordxNo = CLMBordxSEQ!CLMBordxNo
        CLMBordxSEQ.update
    End If
    
    GenerateCLMBordxNoVOID = "KIV" & CMLBordxNo
    CLMBordxSEQ.Close
    Set CLMBordxSEQ = Nothing
End Function







''' Worksheet  Debit & Credit Notes

Function GenerateDebitNo(Payor As String) As String
    Dim DRQ As New ADODB.Recordset
    Dim strSql As String
    Dim SeqNo As String
    
    strSql = "select DRIndex , DRSEQ , Payor from DRSeq where Payor ='" & Payor & "'"
    DRQ.Open strSql, MediConnCISWS, adOpenStatic, adLockOptimistic
    
    If DRQ.RecordCount = 0 Then
        With DRQ
            .AddNew
            !DRSeq = "00001"
            !Payor = Payor
            .update
            SeqNo = "00001"
        End With
    Else
        With DRQ
            !DRSeq = Format(!DRSeq + 1, "00000")
            SeqNo = !DRSeq
            .update
        End With
    End If

    GenerateDebitNo = "D" & Payor & Format(SeqNo, "00000")
    DRQ.Close
    Set DRQ = Nothing
End Function

Function GenerateCreditNo(Payor As String) As String
    Dim CRQ As New ADODB.Recordset
    Dim strSql As String
    Dim SeqNo As String
    
    strSql = "select CRIndex , CRSEQ , Payor from CRSeq where Payor ='" & Payor & "'"
    CRQ.Open strSql, MediConnCISWS, adOpenStatic, adLockOptimistic
    
    If CRQ.RecordCount = 0 Then
        With CRQ
            .AddNew
            !CRSEQ = "00001"
            !Payor = Payor
            .update
            SeqNo = "00001"
        End With
    Else
        With CRQ
            !CRSEQ = Format(!CRSEQ + 1, "00000")
            SeqNo = !CRSEQ
            .update
        End With
    End If

    GenerateCreditNo = "C" & Payor & Format(SeqNo, "00000")
    CRQ.Close
    Set CRQ = Nothing
End Function

Public Function AccessFeeRefund(ExpiryDate As Date, CancelDate As Date, EffDate As Date, MBMAccessFees As Double) As Double
    
    AccessFeeRefund = (CInt(ExpiryDate - CancelDate) / (ExpiryDate - EffDate)) * MBMAccessFees
    If AccessFeeRefund < 0 Then
        AccessFeeRefund = 0
    End If
End Function


Public Function isTime(str As String) As Boolean
    Dim Correct As Boolean
    isTime = False
    Correct = False
    If IsNumeric(Left(str, 2)) Then
        If Left(str, 2) <= 24 Then
            Correct = True
        End If
    End If
    If Correct = True Then
        Correct = False
        If Right(str, 2) < 60 Then
            Correct = True
        End If
    End If
    If Correct = True Then
        isTime = True
    End If
End Function

Public Function GenerateSPVNo(TempBank As String, TmpPVCode)
Dim PVQRec As New ADODB.Recordset
Dim strSql As String
Dim SeqNo As String
Dim PVYear As String
    
    strSql = "Select * from PVSeq2013 where BankCode='" & Trim(TempBank) & "'"
    PVQRec.Open strSql, MediConnCISPayment, adOpenStatic, adLockOptimistic
   
    If PVQRec.RecordCount = 0 Then
        With PVQRec
            .AddNew
            !BankCode = Trim(TempBank)
            !PVCode = Trim(TmpPVCode)
            !PVSEQNO = "100000001"
            .update
            SeqNo = "100000001"
        End With
    Else
        PVQRec!PVSEQNO = Format(PVQRec!PVSEQNO + 1, "000000000")
        SeqNo = PVQRec!PVSEQNO
        PVQRec.update
    End If
    
    GenerateSPVNo = Trim(TmpPVCode) & SeqNo
    PVQRec.Close
    Set PVQRec = Nothing
End Function

Public Function GenerateNPVNo(TempBank As String)
Dim PVQ As New ADODB.Recordset
Dim strSql As String
Dim SeqNo As String
Dim PVYear As String
    
    strSql = "Select * from PVNSeq2013 where BankCode='" & Trim(TempBank) & "'"
    PVQ.Open strSql, MediConnCISPayment, adOpenStatic, adLockOptimistic
    
    If PVQ.RecordCount = 0 Then
        With PVQ
            .AddNew
            !BankCode = Trim(TempBank)
            !PVSEQNO = "100000001"
            .update
            SeqNo = "100000001"
        End With
    Else
        PVQ!PVSEQNO = Format(PVQ!PVSEQNO + 1, "000000000")
        SeqNo = PVQ!PVSEQNO
        PVQ.update
    End If
    
    GenerateNPVNo = Trim(TempBank) & SeqNo
    PVQ.Close
    Set PVQ = Nothing
End Function
'New code for SJ
Public Function GenerateSPVNoforSJ(TempBank As String, TmpPVCode)
Dim PVQRec As New ADODB.Recordset
Dim strSql As String
Dim SeqNo As String
Dim PVYear As String
    strSql = "Select * from PVSeq2013 where BankCode='SJ' and PVCode='" & Trim(TmpPVCode) & "'"
    PVQRec.Open strSql, MediConnCISPayment, adOpenStatic, adLockOptimistic
   
    If PVQRec.RecordCount = 0 Then
        With PVQRec
            .AddNew
            !BankCode = Trim(TempBank)
            !PVCode = Trim(TmpPVCode)
            !PVSEQNO = "100000001"
            .update
            SeqNo = "100000001"
        End With
    Else
        PVQRec!PVSEQNO = Format(PVQRec!PVSEQNO + 1, "000000000")
        SeqNo = PVQRec!PVSEQNO
        PVQRec.update
    End If
    
    GenerateSPVNoforSJ = Trim(TmpPVCode) & SeqNo
    PVQRec.Close
    Set PVQRec = Nothing
End Function


Public Function GenerateGRPClnSeq(GRPCln)
Dim CLNSeq As New ADODB.Recordset
Dim strSql As String
Dim GRPClnCode As String

strSql = "Select * from GRPClnSeq where GRPClnCode = '" & Trim(GRPCln) & "'"
CLNSeq.Open strSql, MediConnCISDB, adOpenStatic, adLockOptimistic

If CLNSeq.RecordCount = 0 Then
    With CLNSeq
        .AddNew
        !GRPClnCode = ChkStr(Trim(GRPCln))
        !GRPClnSeq = "0001"
        .update
        GRPClnCode = ChkStr(Trim(GRPCln)) & "0001"
        GenerateGRPClnSeq = GRPClnCode
    End With
Else
    CLNSeq!GRPClnSeq = Format(CLNSeq!GRPClnSeq + 1, "0000")
    CLNSeq!GRPClnCode = ChkStr(Trim(GRPCln))
    GRPClnCode = CLNSeq!GRPClnSeq
    CLNSeq.update
    GenerateGRPClnSeq = ChkStr(Trim(GRPCln)) & GRPClnCode
End If

CLNSeq.Close
Set CLNSeq = Nothing

End Function

Public Function GenerateGRPCoSeq(GRPCompany)
Dim COMSeq As New ADODB.Recordset
Dim strSql As String
Dim GRPCoCode As String

strSql = "Select * from GrpCoSeq where GrpCoCode = '" & Trim(GRPCompany) & "'"
COMSeq.Open strSql, MediConnCISDB, adOpenDynamic, adLockBatchOptimistic ' adOpenStatic, adLockOptimistic

If COMSeq.EOF = True Then
'If COMSeq.RecordCount = 0 Then
    With COMSeq
        .AddNew
        !GRPCoCode = ChkStr(Trim(GRPCompany))
        !GRPCoSeq = "001"
        .update
        GRPCoCode = ChkStr(Trim(GRPCompany)) & "001"
        GenerateGRPCoSeq = GRPCoCode
    End With
Else
    COMSeq!GRPCoSeq = Format(COMSeq!GRPCoSeq + 1, "000")
    COMSeq!GRPCoCode = ChkStr(Trim(GRPCompany))
    GRPCoCode = COMSeq!GRPCoSeq
    COMSeq.update
    GenerateGRPCoSeq = ChkStr(Trim(GRPCompany)) & GRPCoCode
End If

COMSeq.Close
Set COMSeq = Nothing

End Function

Public Function GeneratePLNSeq(PLNCode)
Dim PLANSeq As New ADODB.Recordset
Dim strSql As String
Dim plancode As String

strSql = "Select * From PLNSeq where PLNSeqCode = '" & Trim(PLNCode) & "'"
PLANSeq.Open strSql, MediConnCISDB, adOpenStatic, adLockOptimistic

If PLANSeq.RecordCount = 0 Then
    With PLANSeq
        .AddNew
        !PLNSeqCode = ChkStr(Trim(PLNCode))
        !PLNNoSeq = "001"
        .update
        plancode = ChkStr(Trim(PLNCode)) & "001"
        GeneratePLNSeq = plancode
    End With
Else
    PLANSeq!PLNNoSeq = Format(PLANSeq!PLNNoSeq + 1, "000")
    PLANSeq!PLNSeqCode = ChkStr(Trim(PLNCode))
    plancode = PLANSeq!PLNNoSeq
    PLANSeq.update
    GeneratePLNSeq = ChkStr(Trim(PLNCode)) & plancode
End If

PLANSeq.Close
Set PLANSeq = Nothing

End Function

Public Function GenerateDocSeq(Doctor)
Dim DRSeq As New ADODB.Recordset
Dim strSql As String
Dim DRCode As String

strSql = "Select * from DocSeq where DocCode = '" & Trim(Doctor) & "'"
DRSeq.Open strSql, MediConnCISDB, adOpenStatic, adLockOptimistic

If DRSeq.RecordCount = 0 Then
    With DRSeq
        .AddNew
        !DocCode = ChkStr(Trim(Doctor))
        !DocSeq = "00001"
        .update
        DRCode = ChkStr(Trim(Doctor)) & "00001"
        GenerateDocSeq = DRCode
    End With
Else
    DRSeq!DocSeq = Format(DRSeq!DocSeq + 1, "00000")
    DRSeq!DocCode = ChkStr(Trim(Doctor))
    DRCode = DRSeq!DocSeq
    DRSeq.update
    GenerateDocSeq = ChkStr(Trim(Doctor)) & DRCode
End If

DRSeq.Close
Set DRSeq = Nothing

End Function

Public Function GenerateCLNSeq(MixClinic)
    Dim CLNSeq As New ADODB.Recordset
    Dim strSql As String
    Dim CLNCode As String
    
    strSql = "Select * from CLNSeq Where CLNCode = '" & Trim(MixClinic) & "'"
    CLNSeq.Open strSql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    
    If CLNSeq.RecordCount = 0 Then
        With CLNSeq
            .AddNew
            !CLNCode = ChkStr(Trim(MixClinic))
            !CLNSeq = "0001"
            .update
            CLNCode = ChkStr(Trim(MixClinic)) & "001"
        End With
    Else
        CLNSeq!CLNSeq = Format(CLNSeq!CLNSeq + 1, "000")
        CLNSeq!CLNCode = ChkStr(Trim(MixClinic))
        CLNCode = CLNSeq!CLNSeq
        CLNSeq.update
    End If

    GenerateCLNSeq = ChkStr(Trim(MixClinic)) & CLNCode
    CLNSeq.Close
    Set CLNSeq = Nothing
End Function

Public Sub MemberSuppInfo(MembershipNo As String, CovPersonID, strBalLimit, strVisitLimit, strAccumAmt, StrDentStatus, strCoPay, strPayor _
As String, strHLTCode As String, strClnPln As String, strClnIns As String, strEffdate As String, strExpDate, strPolNo, strMBMDataStatus, strPOlStatus, strPatientName _
As String, strPatientDOB, strPatientIC, strMBMCLNLTM, strCardType, strVisitIND, strAnnualInd _
As String, strSpecStatus, strMEPStatus, strDentLmt, strSpecLmt, strMEPLmt, SuppLmtType _
As String, StrDentLmtStatus, StrSpecLmtStatus, strBNFStatus, strDentExtLmt, str1stSpecLmt _
As String, strGRPMCOStatus As String, strPLNGrossStatus As String, strMBMDepartment _
As String, strMCStatus As String, strMBMMCDaysBal As String, strGRPPostageFees As String, strGRPBordxFees As String, strPLNMMPerc As String, Strrelationtype1 As String, strMBMCancelDate As String, strCMBMCancelDate As String, strCMBMCCLNExpDate As String)
Dim MBMRec As New ADODB.Recordset
Dim CLNOth As New ADODB.Recordset
Dim MBMSql, CLNSql, COVSql, COVLmtSql As String
Dim COVRec As New ADODB.Recordset
Dim Cnt As Integer
Dim CISPln As New ADODB.Recordset
Dim CISAnnLMt As New ADODB.Recordset
Dim CISPlnSql, CISAnnLmtSql As String
Dim COVLmt As New ADODB.Recordset
Dim CoPayment As New ADODB.Recordset
Dim CoSql As String
Dim tmpMBMNo As String
Dim TempCovID As String
Dim tmpVISLmtIndSupp As String
Dim tmpAnnLmtIndSupp As String
Dim grpComp As New ADODB.Recordset
Dim strsqlGrpComp As String
Dim tmp As String

    SuppLmtType = "A"
    If Len(Trim(MembershipNo)) Then
        MBMSql = "SELECT Top 100 MBMIcBcPp, MBMNumber, MBMName, MBMICBcPp, PAYCode,  MBMCLNStatus, MBMCOVID, MBMGroupCompany, " & _
                " MBMDOB, MBMCLNDataStatus, HLTCode,MBMDepartment,MBMCLNMCDaysBal FROM VIEW_MBM_Clinical WHERE MBMNumber='" & Trim(MembershipNo) & "' AND MBMCLNDataStatus <> 'D'"
                
                '" AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')"
        MBMRec.Open MBMSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        
        If MBMRec.EOF = True Then
            MsgBox "Membership No not found ! ", vbOKOnly + vbCritical, DialogBoxTitle
            frmSearchMBM.Show
           
            tmpMBMNo = MembershipNo & "-" & CovPersonID
            MembershipNo = ""
            CovPersonID = ""
        Else
            tmp = IIf(Len(MBMRec!MBMDepartment), MBMRec!MBMDepartment, "")
            strsqlGrpComp = "SELECT Top 100 GRPMCOStatus, GrpCoName,GRPMCStatus,GRPPostageFees,GRPBordxFees FROM GRPCompany WHERE GrpCoName = '" & Replace(MBMRec!MBMGroupCompany, "'", "''") & "'"
            grpComp.Open strsqlGrpComp, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                
                If grpComp.EOF = False Then
                    strGRPMCOStatus = IIf(Len(grpComp!GRPMCOStatus), grpComp!GRPMCOStatus, "")
                    strMCStatus = IIf(Len(grpComp!GRPMCStatus), grpComp!GRPMCStatus, "")
                    strGRPPostageFees = IIf(Len(grpComp!GRPPostageFees), grpComp!GRPPostageFees, "")
                    strGRPBordxFees = IIf(Len(grpComp!GRPBordxFees), grpComp!GRPBordxFees, "")
                Else
                    strGRPMCOStatus = ""
                    strMCStatus = ""
                    strGRPBordxFees = ""
                End If
            grpComp.Close
            Set grpComp = Nothing
            
            CLNSql = "SELECT Top 100 MBMNumber, MBMPolNumber, MBMPolEffDate, MBMPolExpDate,MBMCLNBalLmt,MBMCLNVisitLmt, MBMAccumAmt, MBMCardType, " & _
                    " INSType, PLNType,MBMCLNLTM, PAYCode, MBMMEPLmt, MBMDentalLmt, MBMSpeclmt, MBMDentalExtLmt, MBMCancelDate, MBMFirstSpecLmt " & _
                    " FROM MBMCLNOthers WHERE MBMNumber='" & Trim(MembershipNo) & "'"
            CLNOth.Open CLNSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            
            CISPlnSql = "SELECT Top 100 AnnLmtIND, VisLmtIND, ClnContract, PDentalStatus, SDentalStatus, " & _
                        " PSpecStatus, SSpecStatus, PMEPStatus, SMEPStatus, BNFStatus, " & _
                        "PDentalLimit, SDentalLimit, PSpecLimit, SSpecLimit, ANNLmtIndSupp, VISLmtIndSupp, GrossAmtStatus, PLNMMIndPerc,ASOInd " & _
                        " from PLN WHERE PLNCode ='" & Trim(CLNOth!PLNType) & "' and HLTCode='" & Trim(MBMRec!HLTCode) & _
                        "' AND PAYCode='" & Trim(CLNOth!PAYCode) & "'"
            CISPln.Open CISPlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            
            CoSql = "SELECT Top 100 * FROM CoPaymentMBM WHERE HLTType='" & MBMRec!HLTCode & "' And PlnCode='" & CLNOth!PLNType & "'"
            CoPayment.Open CoSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If CISPln.EOF = True Then
                MsgBox "Clinical Plan not found ! ", vbCritical + vbOKOnly
                MembershipNo = ""
                CovPersonID = ""
            Else
            tmpVISLmtIndSupp = IIf(Len(Trim(CISPln!VISLmtIndSupp)), Trim(CISPln!VISLmtIndSupp), "Y")
            tmpAnnLmtIndSupp = IIf(Len(Trim(CISPln!ANNLmtIndSupp)), Trim(CISPln!ANNLmtIndSupp), "Y")

            If Trim(CISPln!ANNLmtIndSupp) = "Y" Then
                CISAnnLmtSql = "SELECT SuppLimitStatus FROM AnnualLimit WHERE PLNCode='" & Trim(CLNOth!PLNType) & "' and HLTCode='" & Trim(MBMRec!HLTCode) & "' AND INSCode='" & Trim(CLNOth!INSType) & "'"
                CISAnnLMt.Open CISAnnLmtSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                If CISAnnLMt.EOF = False Then
                    SuppLmtType = CISAnnLMt!SuppLimitStatus
                End If
                CISAnnLMt.Close
                Set CISAnnLMt = Nothing
            End If
            strVisitIND = IIf(Len(CISPln!VISLmtInd), CISPln!VISLmtInd, "")
            If CovPersonID = "00" And SuppLmtType = "D" Then
                strAnnualInd = "N"
            Else
                strAnnualInd = IIf(Len(CISPln!AnnLmtIND), CISPln!AnnLmtIND, "")
            End If
            strBNFStatus = IIf(Len(CISPln!BNFStatus), CISPln!BNFStatus, "")
            'ASOInd = IIf(Len(CISPln!ASOInd), CISPln!ASOInd, "N")
                
            If CovPersonID = "00" And Trim(SuppLmtType) = "D" Then
                strBalLimit = "FULL"
            Else
                strBalLimit = IIf(Trim(CISPln!AnnLmtIND) = "N", "FULL", IIf(Len(CLNOth!MBMCLNBalLmt), CLNOth!MBMCLNBalLmt, 0))
            End If
                
            strVisitLimit = IIf(Trim(CISPln!VISLmtInd) = "N", "FULL", CLNOth!MBMCLNVisitLmt)
            strPLNGrossStatus = IIf(Len(CISPln!GrossAmtStatus), CISPln!GrossAmtStatus, "")
            strPLNMMPerc = IIf(Len(CISPln!PLNMMIndPerc), CISPln!PLNMMIndPerc, 0)
                
            If Trim(CovPersonID) = Trim("00") Then
                StrDentStatus = IIf(Len(CISPln!PDentalStatus), CISPln!PDentalStatus, "")
                strSpecStatus = IIf(Len(CISPln!PSpecStatus), CISPln!PSpecStatus, "")
                strMEPStatus = IIf(Len(CISPln!PMEPStatus), CISPln!PMEPStatus, "")
                StrDentLmtStatus = IIf(Len(CISPln!PDentalLimit), CISPln!PDentalLimit, "")
                StrSpecLmtStatus = IIf(Len(CISPln!PSpecLimit), CISPln!PSpecLimit, "")
            Else
                StrDentStatus = IIf(Len(CISPln!SDentalStatus), CISPln!SDentalStatus, "")
                strSpecStatus = IIf(Len(CISPln!SSpecStatus), CISPln!SSpecStatus, "")
                strMEPStatus = IIf(Len(CISPln!SMEPStatus), CISPln!SMEPStatus, "")
                StrDentLmtStatus = IIf(Len(CISPln!SDentalLimit), CISPln!SDentalLimit, "")
                StrSpecLmtStatus = IIf(Len(CISPln!SSpecLimit), CISPln!SSpecLimit, "")
            End If
        End If
        CISPln.Close
        Set CISPln = Nothing
            
        If CoPayment.EOF = False Then
            strCoPay = CoPayment!CoAmount
        End If
        CoPayment.Close
        Set CoPayment = Nothing
           
        strPayor = CLNOth!PAYCode
        strHLTCode = MBMRec!HLTCode
        strClnPln = CLNOth!PLNType
        strClnIns = CLNOth!INSType
        strEffdate = CLNOth!MBMPolEffDate
        strExpDate = CLNOth!MBMPolExpDate
        strPolNo = IIf(Len(CLNOth!MBMPolNumber), CLNOth!MBMPolNumber, "")
        strMBMDataStatus = MBMRec!MBMCLNDataStatus
        strPOlStatus = MBMRec!MBMCLNStatus
        strPatientName = MBMRec!MBMName
        strPatientDOB = MBMRec!MBMDOB
        strPatientIC = MBMRec!MBMIcBcPp
        strMBMDepartment = IIf(Len(MBMRec!MBMDepartment), MBMRec!MBMDepartment, "")
        strMBMCLNLTM = IIf(Len(Trim(CLNOth!MBMCLNLTM)), CLNOth!MBMCLNLTM, "")
        strDentLmt = IIf(Len(Trim(CLNOth!MBMDentalLmt)), CLNOth!MBMDentalLmt, "")
        strSpecLmt = IIf(Len(Trim(CLNOth!MBMSpecLmt)), CLNOth!MBMSpecLmt, "")
        strMEPLmt = IIf(Len(Trim(CLNOth!MBMMEPLmt)), CLNOth!MBMMEPLmt, "")
        strDentExtLmt = IIf(Len(Trim(CLNOth!MBMDentalExtLmt)), CLNOth!MBMDentalExtLmt, "")
        str1stSpecLmt = IIf(Len(Trim(CLNOth!MBMFirstSpecLmt)), CLNOth!MBMFirstSpecLmt, "")
        strAccumAmt = IIf(IsNumeric(CLNOth!MBMAccumAmt), CLNOth!MBMAccumAmt, 0)
        strCardType = IIf(Len(CLNOth!MBMCardType), CLNOth!MBMCardType, "")
        strMBMMCDaysBal = IIf(IsNumeric(MBMRec!MBMCLNMCDaysBal), MBMRec!MBMCLNMCDaysBal, 0)
        strMBMCancelDate = IIf(Len(CLNOth!MBMCancelDate), CLNOth!MBMCancelDate, "")
        Strrelationtype1 = ""
        If CovPersonID <> "00" And Trim(CovPersonID) <> "" Then
            COVSql = "SELECT Top 100 MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB,ltrim(rtrim(MBMCRELCode)) AS MBMCRELCode FROM MBMCoveredPersons" & _
                     " WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(CovPersonID) & "'"
            COVRec.Open COVSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            If COVRec.EOF = True Then
                MsgBox "No such Cover Person ! ", vbOKOnly + vbCritical, DialogBoxTitle
                MembershipNo = ""
                CovPersonID = ""
                Strrelationtype1 = ""
            Else
                strPatientName = COVRec!MBMCName
                strPatientDOB = COVRec!MBMCDOB
                strPatientIC = COVRec!MBMCICBCPPNo
                Strrelationtype1 = IIf(Len(COVRec!MBMCRELCode), COVRec!MBMCRELCode, "")
            End If
            COVRec.Close
            Set COVRec = Nothing
                
            TempCovID = Trim(CovPersonID) ' supp separate limit type 'C'
            If SuppLmtType = "B" Or SuppLmtType = "D" Then     ' supp shared limit type 'B' And 'D'
                TempCovID = "01"
            End If
            
            COVLmtSql = "SELECT Top 100 MBMCCLNBalLmt, MBMCCLNVisitLmt, MBMCCLNDentalLmt, MBMCCLNSpecLmt, MBMCCLNMEPLmt, " & _
                        "MBMCCLNDentalExtLmt, MBMCCLNFirstSpecLmt, MBMCAccumAmt, MBMCCardType,MBMCoveredPersonsCLN.MBMCCancelDate , MBMCCLNExpDate,MBMCCLNEffDate ,MBMCCLNStatus " & _
                        "FROM MBMCoveredPersonsCLN WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(TempCovID) & "'"
            COVLmt.Open COVLmtSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            If COVLmt.EOF = False Then
                If SuppLmtType = "B" Or SuppLmtType = "C" Or SuppLmtType = "D" Then
                    strBalLimit = IIf(IsNumeric(COVLmt!MBMCCLNBalLmt), COVLmt!MBMCCLNBalLmt, 0)
                    strVisitLimit = IIf(IsNumeric(COVLmt!MBMCCLNVisitLmt), COVLmt!MBMCCLNVisitLmt, 0)
                    If Trim(tmpAnnLmtIndSupp) = "N" Then
                        strBalLimit = "FULL"
                    End If
                    If tmpVISLmtIndSupp = "N" Then
                        strVisitLimit = "FULL"
                    End If
                    strDentLmt = IIf(Len(Trim(COVLmt!MBMCCLNDentalLmt)), COVLmt!MBMCCLNDentalLmt, "")
                    strSpecLmt = IIf(Len(Trim(COVLmt!MBMCCLNSpecLmt)), COVLmt!MBMCCLNSpecLmt, "")
                    strMEPLmt = IIf(Len(Trim(COVLmt!MBMCCLNMEPLmt)), COVLmt!MBMCCLNMEPLmt, "")
                    strDentExtLmt = IIf(Len(Trim(COVLmt!MBMCCLNDentalExtLmt)), COVLmt!MBMCCLNDentalExtLmt, "")
                    str1stSpecLmt = IIf(Len(Trim(COVLmt!MBMCCLNFirstSpecLmt)), COVLmt!MBMCCLNFirstSpecLmt, "")
                    strAccumAmt = IIf(IsNumeric(COVLmt!MBMCAccumAmt), COVLmt!MBMCAccumAmt, 0)
                End If
                strCardType = IIf(Len(COVLmt!MBMCCardType), COVLmt!MBMCCardType, "")
                strMBMDataStatus = COVLmt!MBMCCLNStatus
                strPOlStatus = COVLmt!MBMCCLNStatus
                strCMBMCancelDate = IIf(Len(COVLmt!MBMCCancelDate), COVLmt!MBMCCancelDate, "")
                strCMBMCCLNExpDate = IIf(Len(COVLmt!MBMCCLNExpDate), COVLmt!MBMCCLNExpDate, "")
            Else
                strMBMDataStatus = "D"
                strPOlStatus = "D"
            End If
            COVLmt.Close
            Set COVLmt = Nothing
            CLNOth.Close
            Set CLNOth = Nothing
        End If
    End If
    MBMRec.Close
    Set MBMRec = Nothing
        
    End If
End Sub


Function GenerateCLNMCOCreditNo()
    Dim CRQ As New ADODB.Recordset
    Dim strSql As String
    Dim SeqNo As String
    
    strSql = "select CLNCRNote from CLNMCOCRNoteSeq where 0=0"
    CRQ.Open strSql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    
    If CRQ.RecordCount = 0 Then
        With CRQ
            .AddNew
            !CLNCRNote = "00001"
            .update
            SeqNo = "00001"
        End With
    Else
        With CRQ
            !CLNCRNote = Format(!CLNCRNote + 1, "00000")
            SeqNo = !CLNCRNote
            .update
        End With
    End If

    GenerateCLNMCOCreditNo = "C" & Format(SeqNo, "00000")
    CRQ.Close
    Set CRQ = Nothing
End Function

Public Sub VerifyDiagnosis(DIagCode As String, DiagDesc As String)
Dim Diag As New ADODB.Recordset
Dim DiagSql As String
    
        DiagSql = " SELECT ReasonCode, ReasonDesc " & _
            " from TreatmentReason where ReasonCode='" & Trim(DIagCode) & "'"
        Diag.Open DiagSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        If Diag.EOF Then
            MsgBox "Invalid Diagnosis code ! Please try again ", vbOKOnly + vbCritical
        Else
            DIagCode = Diag!ReasonCode
            DiagDesc = Diag!ReasonDesc
        End If
        Diag.Close
        Set Diag = Nothing
End Sub
Public Function ValidateMBM(MBMNumber As String, LGdate As Date) As Boolean

Dim RecordArray() As String
Dim MBMNo As String
Dim coverid As String
Dim CLNMBMVerify As New ADODB.Recordset
Dim MBMVSql As String

MBMNo = ""
coverid = ""
Erase RecordArray
ValidateMBM = False

RecordArray() = Split(MBMNumber, "-")
MBMNo = RecordArray(0)
coverid = RecordArray(1)

If Len(Trim(MBMNumber)) = 0 Then
    MsgBox "Invalid Membership No., Invalid Membership No.", vbOKOnly + vbCritical, DialogBoxTitle
    ValidateMBM = False
End If

MBMVSql = ""
MBMVSql = "SELECT  MBMNumber, MBMPolEffDate, MBMPolExpDate, MBMCLNStatus, MBMCLNEndtEffDate " & _
        " FROM MBMCLNOthers " & _
        " WHERE MBMNumber ='" & Trim(MBMNo) & "'"
        
If Trim(coverid) <> "00" Then
    MBMVSql = "SELECT  dbo.MBMCLNOthers.MBMNumber, dbo.MBMCLNOthers.MBMPolEffDate, dbo.MBMCLNOthers.MBMPolExpDate, dbo.MBMCLNOthers.MBMCLNStatus, " & _
    "  dbo.MBMCLNOthers.MBMCLNEndtEffDate, dbo.MBMCoveredPersonsCLN.MBMCCoverID, dbo.MBMCoveredPersonsCLN.MBMCCLNStatus " & _
    " FROM dbo.MBMCLNOthers INNER JOIN  dbo.MBMCoveredPersonsCLN ON dbo.MBMCLNOthers.MBMNumber = dbo.MBMCoveredPersonsCLN.MBMCNumber " & _
    " WHERE (dbo.MBMCLNOthers.MBMNumber  ='" & Trim(MBMNo) & "') AND (dbo.MBMCoveredPersonsCLN.MBMCCoverID = '" & Trim(coverid) & "') "
End If
CLNMBMVerify.Open MBMVSql, medixmasterconn, adOpenKeyset, adLockOptimistic

If CLNMBMVerify.RecordCount > 0 Then

        ValidateMBM = True
        
        If (IsDate(LGdate) = False) And (ValidateMBM = True) Then
            MsgBox "Invalid Treatment Date Format, Please Re-enter the Treatment Date.", vbOKOnly + vbCritical, DialogBoxTitle
            ValidateMBM = False
        End If
        
        If (IsDate(CLNMBMVerify!MBMPolEffDate) = False) And (ValidateMBM = True) Then
            MsgBox "Invalid Effective Date Format, Please Check Member Effective Date.", vbOKOnly + vbCritical, DialogBoxTitle
            ValidateMBM = False
        End If
        
        If (IsDate(CLNMBMVerify!MBMPolExpDate) = False) And (ValidateMBM = True) Then
            MsgBox "Invalid Expiry Date Format, Please Check Member Expiry Date.", vbOKOnly + vbCritical, DialogBoxTitle
            ValidateMBM = False
        End If
        
        If (CDate(LGdate) < CDate(CLNMBMVerify!MBMPolEffDate)) And (ValidateMBM = True) Then
            MsgBox "Treatment Date Cannot Less Than Member Effective Date, Please Check Member Effective Date and Treatment Date.", vbOKOnly + vbCritical, DialogBoxTitle
            ValidateMBM = False
        End If
        
        If (CDate(LGdate) > CDate(CLNMBMVerify!MBMPolExpDate)) And (ValidateMBM = True) Then
            MsgBox "Treatment Date Cannot Greater Than Member Expiry Date, Please Check Member Expiry Date and Treatment Date.", vbOKOnly + vbCritical, DialogBoxTitle
            ValidateMBM = False
        End If
        
        If (Trim(CLNMBMVerify!MBMCLNStatus) = "C") And (ValidateMBM = True) Then
            If (IsDate(CLNMBMVerify!MBMCLNEndtEffDate) = False) And (ValidateMBM = True) Then
                MsgBox "Invalid Member Cancellation Date Format, Please Check Member Cancel Date.", vbOKOnly + vbCritical, DialogBoxTitle
                ValidateMBM = False
            End If
            
            If (CDate(LGdate) >= CDate(CLNMBMVerify!MBMCLNEndtEffDate)) And (ValidateMBM = True) Then
                MsgBox "Membership Has Been Canceled , Please Check Membership.", vbOKOnly + vbCritical, DialogBoxTitle
                ValidateMBM = False
            End If
        End If
        
        If (Trim(coverid) <> "00") And (ValidateMBM = True) Then
            If (Trim(CLNMBMVerify!MBMCLNStatus) = "A") And (Trim(CLNMBMVerify!MBMCCLNStatus) = "C") And (ValidateMBM = True) Then
                MsgBox "Membership Has Been Canceled , Please Check Membership.", vbOKOnly + vbCritical, DialogBoxTitle
                ValidateMBM = False
            End If
        End If
Else
    MsgBox "Membership Validation Failed! ", vbCritical
    ValidateMBM = False
End If


End Function

Public Function GetLGCode(MBMNumber As String, Ind As String, Optional LGNo As String)
Dim LG As New ADODB.Recordset
Dim SEQ As New ADODB.Recordset
Dim strSql As String
Dim strseq As String
Dim LGSeq As String
Dim LGNumber As String
     '& "' order by AnnEffDate  desc"
    strSql = "Select * from LGSeq Where LGNo = '" & Trim(Mid(LGNo, 3, 5)) & "'  order by LGNo Desc"
    'strseq = ""
    strseq = "Select * From LGNo  order by LGNo Desc"
    LG.Open strSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    SEQ.Open strseq, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If Ind = "N" Then
        If SEQ.RecordCount = 0 Then
            With SEQ
                .AddNew
                    !MBMNumber = Trim(Mid(MBMNumber, 1, InStr(MBMNumber, "-") - 1))
                    !LGNo = Format(1, "00000")
                    LGNumber = !LGNo
                .update
            End With
        Else
            'SEQ.MoveLast
            LGNumber = SEQ!LGNo
            With SEQ
                .AddNew
                    !MBMNumber = Trim(Mid(MBMNumber, 1, InStr(MBMNumber, "-") - 1))
                    LGNumber = Format(LGNumber + 1, "00000")
                    !LGNo = LGNumber
                .update
            End With
        End If
    Else
        LGNumber = Mid(LGNo, 3, 5)
    End If
    
    If LG.RecordCount = 0 Then
        With LG
            .AddNew
                !LGNo = LGNumber
                !LGSequence = "1"
                LGSeq = !LGSequence
            .update
        End With
    Else
        With LG
                LGSeq = Format(!LGSequence + 1)
                !LGSequence = LGSeq
                !LGNo = LGNumber
        .update
        End With
    End If
    'GenerateDebitNo = "D" & payor & SEQNO
    GetLGCode = "LG" & LGNumber & "-" & LGSeq
    LG.Close
    Set LG = Nothing
End Function

Public Function GenSUNPostingBatchNo()
Dim SeqRec As New ADODB.Recordset
Dim strSql As String
Dim NewSeqCode As String
Dim NewRec As Boolean

    SeqRec.Open "SunPostBatch", MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    NewRec = True
    NewSeqCode = "1001"
    Do While Not SeqRec.EOF
        NewRec = False
        NewSeqCode = SeqRec!BatchNo + 1
        SeqRec.MoveNext
    Loop
    SeqRec.Close
    Set SeqRec = Nothing
    If NewRec = True Then
        strSql = "Insert Into SunPostBatch(BatchNo) values('" & NewSeqCode & "')"
    Else
        strSql = "Update SunPostBatch SET BatchNo = '" & NewSeqCode & "'"
    End If
    MediConnCISWS.Execute strSql
    GenSUNPostingBatchNo = Trim(NewSeqCode)
End Function

Public Function GenerateMembershipNo(PayorCode, InsuredType)
    Dim SEQ As New ADODB.Recordset
    Dim SEQInsert As New ADODB.Recordset
    Dim currentNo, i As Integer
    Dim strSql As String
    Dim SEQNumber, SEQCode As String
    Dim CharacterArray
    Dim firstChar, currentChar As String
    

    
    SEQNumber = "00001"
    CharacterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType) & Trim(BusinessYear))
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType))
    SEQCode = Trim(PayorCode & Trim(InsuredType))
    
    strSql = "Select * from MBMSEQ where SEQCOde='" & Trim(SEQCode) & "' "
    SEQ.Open strSql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    currentNo = 1
    currentChar = "A"
    
    If SEQ.RecordCount = 0 Then
        SEQInsert.Open "Select * from MBMSEQ ", medixmasterconnMaint, adOpenStatic, adLockOptimistic
        With SEQInsert
            .AddNew
            !SEQCode = Trim(SEQCode)
            !PAYCode = PayorCode
            '!PLNCode = PlanCode
            !INSCode = InsuredType
            !SEQNumber = currentChar & SEQNumber
            SEQNumber = currentChar & SEQNumber
            .update
        End With
        SEQInsert.Close
        Set SEQInsert = Nothing
    Else
        currentNo = Mid(SEQ!SEQNumber, 2, 5)
        firstChar = Mid(SEQ!SEQNumber, 1, 1)
        Do Until i = 24
            If firstChar = CharacterArray(i) Then
                If currentNo <> "99999" Then
                    currentChar = CharacterArray(i)
                    currentNo = currentNo + 1
                Else
                    currentChar = CharacterArray(i + 1)
                    currentNo = "00001"
                End If
                Exit Do
            End If
            i = i + 1
        Loop
        SEQNumber = currentChar & Format(currentNo, "00000")
        SEQ!SEQNumber = SEQNumber
        SEQ.update
    End If

    GenerateMembershipNo = SEQNumber
    SEQ.Close
    Set SEQ = Nothing
End Function

Public Function GeneratePostingINVNo()
    Dim SEQ As New ADODB.Recordset
    Dim SEQInsert As New ADODB.Recordset
    Dim currentNo, i As Integer
    Dim strSql As String
    Dim SEQNumber, SEQCode As String
    Dim CharacterArray
    Dim firstChar, currentChar As String
    

    
    SEQNumber = "000001"
    CharacterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType) & Trim(BusinessYear))
    'SEQCode = Trim(PayorCode & Trim(PlanCode) & Trim(InsuredType))
    'SEQCode = Trim(PayorCode & Trim(InsuredType))
    
    'StrSql = "Select * from PostingINVSEQ where SEQCOde='" & Trim(SEQCode) & "' "
    strSql = "Select * from PostingINVSEQ "
    SEQ.Open strSql, MediConnCISWS, adOpenStatic, adLockOptimistic
    currentNo = 1
    currentChar = "A"
    
    If SEQ.RecordCount = 0 Then
        'SEQInsert.Open "Select * from MBMSEQ ", medixmasterconnMaint, adOpenStatic, adLockOptimistic
        With SEQ
            .AddNew
            '!SEQCode = Trim(SEQCode)
            '!PAYCode = PayorCode
            '!PLNCode = PlanCode
            '!INSCode = InsuredType
            !SEQNumber = currentChar & SEQNumber
            SEQNumber = currentChar & SEQNumber
            .update
        End With
        SEQ.Close
        Set SEQ = Nothing
    Else
        currentNo = Mid(SEQ!SEQNumber, 2, 6)
        firstChar = Mid(SEQ!SEQNumber, 1, 1)
        Do Until i = 24
            If firstChar = CharacterArray(i) Then
                If currentNo <> "999999" Then
                    currentChar = CharacterArray(i)
                    currentNo = currentNo + 1
                Else
                    currentChar = CharacterArray(i + 1)
                    currentNo = "000001"
                End If
                Exit Do
            End If
            i = i + 1
        Loop
        SEQNumber = currentChar & Format(currentNo, "000000")
        SEQ!SEQNumber = SEQNumber
        SEQ.update
        SEQ.Close
        Set SEQ = Nothing
    End If

    GeneratePostingINVNo = SEQNumber

End Function

