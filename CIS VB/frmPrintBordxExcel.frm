VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintBordxExcel 
   Caption         =   "Print Bordx (Excel)"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Width           =   11775
      Begin VB.ComboBox cboBilltype 
         Height          =   315
         ItemData        =   "frmPrintBordxExcel.frx":0000
         Left            =   9240
         List            =   "frmPrintBordxExcel.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   195
         Width           =   1575
      End
      Begin VB.ComboBox CboReportType 
         Height          =   315
         ItemData        =   "frmPrintBordxExcel.frx":002A
         Left            =   6360
         List            =   "frmPrintBordxExcel.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   200
         Width           =   2055
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         ItemData        =   "frmPrintBordxExcel.frx":00B1
         Left            =   1800
         List            =   "frmPrintBordxExcel.frx":00B3
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ComboBox CboMonth 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrintBordxExcel.frx":00B5
         Left            =   4440
         List            =   "frmPrintBordxExcel.frx":00DD
         TabIndex        =   20
         Top             =   675
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox CboYear 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrintBordxExcel.frx":0143
         Left            =   6720
         List            =   "frmPrintBordxExcel.frx":0145
         TabIndex        =   19
         Top             =   675
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   200
         Width           =   1575
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   3600
         TabIndex        =   18
         Top             =   200
         Width           =   1455
      End
      Begin VB.Label lblBilltype 
         Caption         =   "Bill Type"
         Height          =   255
         Left            =   8520
         TabIndex        =   32
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Report Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   28
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Month"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblDept 
         Caption         =   "Department"
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Company"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   22
         Top             =   200
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   7200
      Width           =   11775
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   480
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   495
         Left            =   9480
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   495
         Left            =   8760
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   10200
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   495
         Left            =   10920
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintSummary 
         Caption         =   "&Print Summary"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6960
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintDetails 
         Caption         =   "Print &Detail"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7920
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSummaryPreview 
         Caption         =   "S&ummary Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6120
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdDetailPreview 
         Caption         =   "De&tail Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtPStatus 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Print Bordx (Excel)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Width           =   3285
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   6135
      Left            =   0
      TabIndex        =   16
      Top             =   1080
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10821
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Plan"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Claim No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patient Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Clinic"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Clinic Name"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Bill Amt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Co-Pay"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "App Amt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "FFS"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "BordxAmt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Print Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "HLT Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ReportType"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPrintBordxExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intExit As Integer
Dim ScanCode As String
Dim RePrintStatus As Boolean
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim tmpPrintStatus As Boolean
Dim tmpRptByDept As Boolean
Dim GRPCompany As String
Dim tmpRptFormat As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim tmpRptType As String
Dim tmpRptThphoidType As String
Dim tmpcovercompany As String
Dim tmpcoveraadd1 As String
Dim tmpcoveraadd2 As String
Dim tmpcoveraadd3 As String
Dim tmpcoveraadd4 As String
Dim tmpcoveraadd5 As String
Dim tmpcoveraadd6 As String
Dim tmpcoverbordxrefno As String
Dim tmpcoverbordxdate As String
Dim tmpbordxtype As String
Dim tmpbordxcoveramt As String
Dim compCodeByBordx As String

Private m_blnDirection(1 To 13) As Boolean


Private Sub cboCompany_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCompany, KeyAscii)
End Sub

Private Sub CboMonth_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboMonth, KeyAscii)
End Sub

Private Sub CboYear_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboYear, KeyAscii)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub
Private Sub ComboKeyPress(tmpCombo As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(tmpCombo.Text, tmpCombo.SelStart) + Chr(KeyAscii) + Mid(tmpCombo.Text, tmpCombo.SelStart + tmpCombo.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To tmpCombo.ListCount - 1
       
        If NewText = LCase(Left(tmpCombo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = tmpCombo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               tmpCombo.Text = ValidValue
               tmpCombo.SelStart = 0
               tmpCombo.SelLength = Len(ValidValue)
            End If
        End If

End Sub
Private Sub DetailPrint()

Const METHOD_NAME = "DetailPrint"
On Error GoTo debugs


Dim StrSql1 As String
Dim strsqlPlutoBor As String
Dim PlutoBor As New ADODB.Recordset
Dim sqlinsert As String
Dim sqlbor As String
Dim tmpbordxdate As String
Dim tmpDateVisit As String
Dim TempMBMName As String
Dim TempMBMDept As String
Dim tempPath As String
Dim PLUTOStatus As Boolean
Dim BOR As New ADODB.Recordset
Dim TempRemarks As String
Dim tempGrpCompany As String
Dim TempCLNName As String
Dim TempMBMCName As String
Dim TempPatientName As String
Dim TempCASDiagnosis5 As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim TempMBMBranch As String
Dim CLM As New ADODB.Recordset
Dim sql As String
Dim sql2 As String
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim CaseCnt As String

    bPrintPass = False
    CancelPrinting = False
    Screen.MousePointer = vbHourglass
    CaseCnt = 0
        sqlbor = " And BordxRefNo = '" & ChkStr(Trim(TxtBordxNo)) & "'"
        
        Set cmd.ActiveConnection = MediConnCISWS
        cmd.CommandText = "ClaimBordxModule"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sqlbor)
        cmd.Parameters.Append Prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "12")
        cmd.Parameters.Append prm2
        BOR.CursorLocation = adUseClient
        BOR.CursorType = adOpenForwardOnly
        BOR.CacheSize = 100
        Set BOR = cmd.Execute
        

                Do While Not BOR.EOF
                    If objExcel Is Nothing Then
                        Set objExcel = CreateObject("Excel.application")
                            objExcel.Interactive = False
                        If objExcel.Visible = False Then
                            objExcel.Visible = True
                        End If
                        
                        objExcel.WindowState = xlMaximized
                        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\IBMOutpatientMedicalClaims.xlt")
                        Set objWorksheet = objWorkbook.Sheets(2)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                    End If
                    
                    tmpbordxdate = Format(BOR!BordxDate, "yyyy/mm/dd")
                    tmpDateVisit = Format(BOR!DateVisit, "yyyy/mm/dd")
                    TempMBMName = Replace(Trim(BOR!MBMName), "'", "''")
                    TempRemarks = IIf(Len(BOR!Remarks), BOR!Remarks, "")
                    tempGrpCompany = Replace(Trim(BOR!MBMGroupCompany), "'", "''")
                    TempMBMDept = Replace(Trim(BOR!MBMDepartment), "'", "''")
                    TempCLNName = Replace(Trim(BOR!CLNName), "'", "''")
                    TempPatientName = Replace(Trim(BOR!PatientName), "'", "''")
                    If Len(BOR!MBMCName) Then
                        TempMBMCName = Replace(Trim(BOR!MBMCName), "'", "''")
                    End If
                    If Len(BOR!CASDiagnosis5) Then
                        TempCASDiagnosis5 = Replace(Trim(BOR!CASDiagnosis5), "'", "''")
                    Else
                        TempCASDiagnosis5 = ""
                    End If
                    
                    If Len(BOR!MBMBranch) Then
                        TempMBMBranch = Trim(BOR!MBMBranch)
                    Else
                        TempMBMBranch = ""
                    End If
                    
                    
                    
                    'sqlinsert = ""
                    'sqlinsert = "Insert into TempBordxSummary(CASNumber, BordxRefNo, CLMTotalActual, CLMCoPayColl," & _
                    '                                            "CLMBordxAmt, CLMFinalApprove,CLMMCoProcFee,MBMnumber," & _
                    '                                            "CLMPayableByMedix2C,CLMPayableByMedix2M, CLMTotalApproved, " & _
                    '                                            "DebitCreditRefNo, CLMWSRemarks,PatientName,Relation, " & _
                    '                                            "CoverID,DateVisit,MBMName,GRPCompany,MBMPolicyNo,PLNCode, " & _
                    '                                            "MBMEmployeeNo,HLTCode,MBMCName,MBMMemberType,PAYCode,INSCode, " & _
                    '                                            "CLNName,CLNCode,MBMCRELCode,RELCode,MBMGroupCompany,MBMDepartment, " & _
                    '                                            "Remarks,BordxPrintedStatus,BordxDate,BordxPrintMonth,BordxPrintYear, " & _
                    '                                            "ReasonDesc1,ReasonDesc2,ReasonDesc3,CASDiagnosis5,MBMCostCenter,GRPBordxFormat, MBMBranch) " & _
                    '            "Values('" & BOR!CasNumber & "', '" & BOR!BordxRefNo & "', " & BOR!CLMTotalActual & ", " & BOR!CLMCoPayColl & "," & _
                    '                                            "" & BOR!CLMBordxAmt & ", " & BOR!CLMFinalApprove & ", " & BOR!CLMMCOProcFee & ", '" & BOR!MBMNumber & "', " & _
                    '                                            "" & IIf(IsNumeric(BOR!CLMPayableByMedix2C), BOR!CLMPayableByMedix2C, "0") & ", " & IIf(IsNumeric(BOR!CLMPayableByMedix2M), BOR!CLMPayableByMedix2M, "0") & ", " & IIf(IsNumeric(BOR!CLMTotalApproved), BOR!CLMTotalApproved, "0") & ", " & _
                    '                                            "'" & BOR!DebitCreditRefNo & "', '" & BOR!CLMWSRemarks & "', '" & TempPatientName & "', '" & BOR!Relation & "', " & _
                    '                                            "'" & BOR!coverid & "','" & tmpDateVisit & "','" & TempMBMName & "','" & BOR!GRPCompany & "','" & BOR!MBMPolicyNo & "','" & BOR!PLNCode & "', " & _
                    '                                            "'" & BOR!MBMEmployeeNo & "','N','" & TempMBMCName & "','" & BOR!MBMMemberType & "','" & BOR!PAYCode & "','" & BOR!INSCode & "', " & _
                    '                                            "'" & TempCLNName & "', '" & BOR!CLNCode & "', '" & BOR!MBMCRELCode & "','" & BOR!RELCODE & "', '" & tempGrpCompany & "','" & TempMBMDept & "', " & _
                    '                                            "'" & TempRemarks & "','" & IIf(Trim(BOR!BordxPrintedStatus) = "TRUE", "1", "0") & "','" & tmpbordxdate & "', '" & BOR!BORDXPrintMonth & "','" & BOR!BORDXPrintYear & "', " & _
                    '                                            "'" & BOR!ReasonDesc1 & "','" & BOR!ReasonDesc2 & "','" & BOR!ReasonDesc3 & "','" & TempCASDiagnosis5 & "', '" & BOR!MBMCostCenter & "', '" & tmpRptFormat & "', '" & TempMBMBranch & "' )"
                    'MediConnCISWSPluto.Execute sqlinsert
               With BOR
                    CaseCnt = CDbl(CaseCnt) + 1
                    
                        If CaseCnt = 1 Then
                            objWorksheet.Cells(1 + CDbl(CaseCnt), "C") = BOR!Bordxrefno
                            objWorksheet.Cells(2 + CDbl(CaseCnt), "C") = tmpbordxdate
                        End If
                        
                        
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "A") = TempMBMDept
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "C") = !CasNumber
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "D") = !MBMEmployeeNo
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "E") = TempMBMName
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "F") = !MBMNumber
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "G") = !coverid
                        If BOR!coverid = "00" Then
                            objWorksheet.Cells(5 + CDbl(CaseCnt), "H") = TempMBMName
                        Else
                            objWorksheet.Cells(5 + CDbl(CaseCnt), "H") = TempPatientName
                        End If
                        'objWorksheet.Cells(5 + CDbl(CaseCnt), "F") = BOR!MBMIcBcPp

                        'If BOR!MBMPatientCovID = "00" Then
                        '    objWorksheet.Cells(5 + CDbl(CaseCnt), "G") = BOR!MBMIcBcPp
                        'Else
                        '    objWorksheet.Cells(5 + CDbl(CaseCnt), "G") = BOR!MBMCICBCPPNo
                        'End If
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "I") = TempCLNName
                        
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "J") = !CLNType
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "K") = !CLMTotalActual
                        'If BOR!CLNType = "G" Then
                            objWorksheet.Cells(5 + CDbl(CaseCnt), "L") = !CLMBordxAmt
                        'Else
                            objWorksheet.Cells(5 + CDbl(CaseCnt), "M") = 0
                        'End If
                        'ReasonDesc1,ReasonDesc2,ReasonDesc3,CASDiagnosis5
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "N") = !CLMBordxAmt
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "O") = BOR!INVPayTo
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "P") = tmpDateVisit
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "Q") = BOR!MCDays
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "R") = !ReasonDesc1
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "S") = !ReasonDesc2
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "T") = !ReasonDesc3
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "U") = !CASDiagnosis5
                        
                End With
                    
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
            
        
                Set objWorksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objWorksheet = objWorkbook.Worksheets.Item(1)
                sqlbor = ""
                CaseCnt = 0
                sqlbor = "Select * from View_CIS_PrintBordx_Summary   where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                BOR.Open sqlbor, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                Do While Not BOR.EOF
                    CaseCnt = CDbl(CaseCnt) + 1
                    If CaseCnt = 1 Then
                        objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                        objWorksheet.Cells(12, "K") = tmpbordxdate
                        objWorksheet.Cells(20, "B") = "Actual Outpatient medical cost incurred for the month of" & " " & CboMonth & " " & CboYear
                    End If
                    
                    
                       
                        If Trim(BOR!MBMDepartment) = "REGULAR" Then
                            'If BOR!CLNType = "G" Then
                                objWorksheet.Cells(24, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(24, "I"))
                                objWorksheet.Cells(24, "K") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(24, "M") = CDbl(BOR!CLMBordxAmt)
                            'Else
                            '    objWorksheet.Cells(24, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(24, "I"))
                            '    objWorksheet.Cells(24, "L") = CDbl(BOR!CLMBordxAmt)
                            'End If
                        ElseIf Trim(BOR!MBMDepartment) = "FTH" Then
                            'If BOR!CLNType = "G" Then
                                objWorksheet.Cells(26, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(26, "I"))
                                objWorksheet.Cells(26, "K") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(26, "M") = CDbl(BOR!CLMBordxAmt)
                            'Else
                            '    objWorksheet.Cells(26, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(26, "I"))
                            '    objWorksheet.Cells(26, "L") = CDbl(BOR!CLMBordxAmt)
                            'End If
                        ElseIf Trim(BOR!MBMDepartment) = "RETIREE" Then
                            'If BOR!CLNType = "G" Then
                                objWorksheet.Cells(28, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(28, "I"))
                                objWorksheet.Cells(28, "K") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(28, "M") = CDbl(BOR!CLMBordxAmt)
                            'Else
                            '    objWorksheet.Cells(28, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(28, "I"))
                            '    objWorksheet.Cells(28, "L") = CDbl(BOR!CLMBordxAmt)
                            'End If
                        Else
                            'If BOR!CLNType = "G" Then
                                objWorksheet.Cells(30, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(30, "I"))
                                objWorksheet.Cells(30, "K") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(30, "M") = CDbl(BOR!CLMBordxAmt)
                            'Else
                            '    objWorksheet.Cells(30, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(30, "I"))
                            '    objWorksheet.Cells(30, "L") = CDbl(BOR!CLMBordxAmt)
                            'End If
                        End If
                            
                        
                    CaseCnt = CDbl(CaseCnt) + 1
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
    'End If
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Set objWorksheet = Nothing
    Set objExcel = Nothing
    Screen.MousePointer = Default
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub cmdPrintDetails_Click()
    If cboCompany = "" Then
        MsgBox "Please Select company name", vbOKOnly + vbCritical, DialogBoxTitle
        cboCompany.SetFocus
    ElseIf cboBilltype = "" Then
        MsgBox "Please Select Bill Type", vbOKOnly + vbCritical, DialogBoxTitle
        cboBilltype.SetFocus
    ElseIf cboCompany = "Sihat Johor" Then
        Call PrintSJTemplate
    ElseIf cboCompany = "Health Connect" Or cboCompany = "MediExpress" Then
        Call PrintTemplate
    Else
        MsgBox "Please Select Health Connect Or Sihat Johor..!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub
Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPrintBordxExcel")
Dim AA As Integer
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    CboYear.Clear
    For AA = 2012 To Year(Date) + 2
        CboYear.AddItem AA
    Next AA
    CmdPrintSummary.Enabled = False
    cmdPrintDetails.Enabled = False
    
    cboCompany.AddItem "MediExpress"
    cboCompany.AddItem "Health Connect"
    cboCompany.AddItem "Sihat Johor"
End Sub

'******************** Start Command Button ********************************
Private Sub Cmdclear_Click()
    lstPrintBordx.ListItems.Clear
    TxtBordxNo.Text = 0
    CboMonth.Text = ""
    CboYear.Text = ""
    TxtBordxNo.Text = ""
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    CmdPrintSummary.Enabled = False
    cmdPrintDetails.Enabled = False
    CmdSummaryPreview.Enabled = False
    cmdDetailPreview.Enabled = False
    CboMonth.Enabled = False
    CboYear.Enabled = False
    TxtBordxNo.SetFocus
End Sub

Private Sub cmdSearch_Click()
Dim strAppend As String

    Screen.MousePointer = vbHourglass
    If Trim(TxtBordxNo) = "" Then
        MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
        TxtBordxNo.SetFocus
    Else
        ScanCode = TxtBordxNo
        If Mid(ScanCode, 1, 2) = "LO" Then
            TxtBordxNo = ""
        Else
            CmdClear.Enabled = False
            strAppend = " CISWorksheet.dbo.Claim.BordxRefNo = '" & ChkStr(Trim(TxtBordxNo)) & "'"
            CmdPrintSummary.Enabled = True
            cmdPrintDetails.Enabled = True
            CmdSummaryPreview.Enabled = True
            cmdDetailPreview.Enabled = True
            Call PrintBordxListing(strAppend)
        End If
    End If
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub PrintBordxListing(Optional strAppend As String)
Const METHOD_NAME = "PrintBordxListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim TotRec As String
Dim CurrRec As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Check As Boolean
Dim CLM As New ADODB.Recordset
Dim sql As String
Dim sql2 As String
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim strAppend2 As String
    tmpRptType = ""
    tmpRptThphoidType = ""
    TotRec = 0
    CurrRec = 0
    Check = False
    lstPrintBordx.ListItems.Clear
    'sql = " Select  CASNumber, BordxRefNo From Claim where BordxRefNo = '" & TxtBordxNo & "'"
    'CLM.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        Set cmd2.ActiveConnection = MediConnCISWS
        cmd2.CommandText = "CLM_PayAdviceCLNBSN"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 4000, strAppend)
        cmd2.Parameters.Append Prm3
        Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "1")
        cmd2.Parameters.Append prm4
        CLM.CursorLocation = adUseClient
        CLM.CursorType = adOpenForwardOnly
        CLM.CacheSize = 100
        Set CLM = cmd2.Execute
    
    If CLM.EOF = False Then
    '    lstPrintBordx.ListItems.Clear
    '    TotalAmt = 0
    '    Do While Not CLM.EOF
            'strAppend2 = ""
            'strAppend2 = " And CasNumber = '" & Trim(CLM!CasNumber) & "'"
            
            'Set cmd.ActiveConnection = MediConnCISWS
            'cmd.CommandText = "ClaimBordxModule"
            'cmd.CommandType = adCmdStoredProc
            'cmd.CommandTimeout = 300
            'Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend2)
            'cmd.Parameters.Append Prm1
            'Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "12")
            'cmd.Parameters.Append Prm2
            'rst2.CursorLocation = adUseClient
            'rst2.CursorType = adOpenForwardOnly
            'rst2.CacheSize = 100
            'Set rst2 = cmd.Execute
            'sql = ""
            'sql = " Select PLNCode, CASNumber, BordxRefNo, MBMGroupCompany,RptFormat, TyphoidRpt," & _
            '      " PatientName, MBMNumber, CLNName, CLMPayableByMedix2C,CLMFinalApprove, GRPBordxFormat," & _
            '      " CLMTotalApproved, CLMPayableByMedix2M, DebitCreditRefNo,CLMCoPayCLNColl, " & _
            '      " BordxPrintedStatus, HLTCode, PAYCode, ClinicCode,CLMTotalActual,CLMMCOProcFee,CLMBordxAmt, " & _
            '      " MBMDepartment From View_CIS_PrintBordx where CasNumber = '" & Trim(CLM!CasNumber) & "'"
            ' rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            'If rst2.EOF = True Then
            '    MsgBox "No record found ", vbOKOnly + vbCritical, DialogBoxTitle
            '    CmdPrintSummary.Enabled = False
            '    cmdPrintDetails.Enabled = False
            '    CmdSummaryPreview.Enabled = False
            '    cmdDetailPreview.Enabled = False
            'Else
                Do
                    Do Until CLM.EOF
                        'TotRec = CLM.RecordCount
                        'lblTotal = TotRec
                         With CLM
                            GRPCompany = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
                            Set Record = lstPrintBordx.ListItems.Add(, , !PLNCode)
                            Record.SubItems(1) = IIf(Len(Trim(!CasNumber)), Trim(!CasNumber), "")
                            If !coverid = "00" Then
                                Record.SubItems(2) = IIf(Len(Trim(!MBMName)), Trim(!MBMName), "")
                            Else
                                Record.SubItems(2) = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                            End If
                            Record.SubItems(3) = IIf(Len(Trim(!MBMNumber)), Trim(!MBMNumber), "")
                            Record.SubItems(4) = IIf(Len(Trim(!ClinicCode)), Trim(!ClinicCode), "")
                            Record.SubItems(5) = IIf(Len(Trim(!CLNName)), Trim(!CLNName), "")
                            Record.SubItems(6) = IIf(Len(!CLMTotalActual), Format(!CLMTotalActual, "#,##0.00"), 0)
                            Record.SubItems(7) = IIf(Len(!CLMCoPayCLNColl), Format(!CLMCoPayCLNColl, "#,##0.00"), 0)
                            Record.SubItems(8) = IIf(Len(!CLMFinalApprove), Format(!CLMFinalApprove, "#,##0.00"), 0)
                            Record.SubItems(9) = IIf(Len(!CLMMCOProcFee), Format(!CLMMCOProcFee, "#,##0.00"), 0)
                            Record.SubItems(10) = IIf(Len(!CLMBordxAmt), Format(!CLMBordxAmt, "#,##0.00"), 0)
                            Record.SubItems(11) = IIf(Len(Trim(!Bordxrefno)), Trim(!Bordxrefno), "")
                            Record.SubItems(12) = "NO"
                            RePrintStatus = !BordxPrintedStatus
                            tmpRptFormat = IIf(Len(Trim(!GRPBordxFormat)), Trim(!GRPBordxFormat), "")
                            If RePrintStatus = True Then
                                Record.SubItems(12) = "YES"
                            End If
                            tmpRptType = IIf(Len(Trim(!RptFormat)), Trim(!RptFormat), "")
                            tmpRptThphoidType = IIf(Len(Trim(!TyphoidRpt)), Trim(!TyphoidRpt), "")
                            CurrRec = CurrRec + 1
                            lblCurrent = CurrRec
                            If IsNumeric(TotalAmt) = False Then
                                TotalAmt = 0
                            End If
                            TotalAmt = CDbl(TotalAmt) + !CLMBordxAmt
                            LblTotalAmt = Format(TotalAmt, "#,##0.00")
                        End With
                        CLM.MoveNext
                        If intExit = 1 Then
                           Check = False
                           Exit Do
                        End If
                    DoEvents
                        If intExit = 1 Then
                           Check = False
                           Exit Do
                        End If
                    Loop
                Loop Until Check = False
            'End If
            intExit = 0
            'CLM.Close
            'Set CLM = Nothing
                        
        'CLM.MoveNext
        'Loop
        'CmdClear.Enabled = True
            
        Else
            MsgBox "No record found ", vbOKOnly + vbCritical, DialogBoxTitle
            CmdPrintSummary.Enabled = False
            cmdPrintDetails.Enabled = False
            CmdSummaryPreview.Enabled = False
            cmdDetailPreview.Enabled = False
        End If
        CLM.Close
        Set CLM = Nothing
        
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub lstPrintBordx_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(13)
    End Select
End Sub

Private Sub CboMonth_Change()
    If InStr(CboMonth.Text, "null") Then
       CboMonth.Enabled = False
    End If
End Sub

Private Sub CboYear_Change()
    If InStr(CboYear.Text, "null") Then
       CboYear.Enabled = False
    End If
End Sub

Private Sub TxtBordxNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtBordxNo_LostFocus()
    TxtBordxNo = ChkStr(TxtBordxNo)
    CmdSearch.Default = False
    If Trim(TxtBordxNo) <> "" Then
        'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            Call GetBordxInfo
        'MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
    End If
End Sub

Private Sub UpdateBordx()
Const METHOD_NAME = "UpdateBordx"
On Error GoTo debugs
Dim strSql As String
    strSql = " update Bordx set " & _
             " BordxPrintedDate = '" & Format(Date, "YYYY/MM/DD") & "', " & _
             " BORDXPrintMonth = '" & ChkStr(Trim(CboMonth)) & "', " & _
             " BORDXPrintYear = '" & Trim(CboYear) & "', " & _
             " BordxPrintedStatus = '1', " & _
             " BordxPrintedBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
    MediConnCISWS.Execute strSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub GetBordxInfo()

Const METHOD_NAME = "UpdateBordx"
On Error GoTo debugs

Dim StrSqlBordx As String
Dim BordxREc As New ADODB.Recordset
Dim tmpGrpCom As String
    
    tmpGrpCom = ""
    tmpRptByDept = 0
    tmpPrintStatus = 0
    CboMonth.Enabled = False
    CboYear.Enabled = False
    StrSqlBordx = "SELECT GroupCompany, RptTypeByDept,  " & _
                  "BordxPrintedStatus, BORDXPrintMonth, BORDXPrintYear " & _
                  " From Bordx Where BordxRefNo='" & Trim(TxtBordxNo) & "'"
    BordxREc.Open StrSqlBordx, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    Do While Not BordxREc.EOF
        CboMonth.Text = IIf(Len(Trim(BordxREc!BORDXPrintMonth)), Trim(BordxREc!BORDXPrintMonth), "")
        CboYear.Text = IIf(Len(Trim(BordxREc!BORDXPrintYear)), Trim(BordxREc!BORDXPrintYear), "")
        tmpGrpCom = IIf(Len(Trim(BordxREc!GroupCompany)), Trim(BordxREc!GroupCompany), "")
        tmpRptByDept = IIf(Len(Trim(BordxREc!RptTypeByDept)), Trim(BordxREc!RptTypeByDept), 0)
        tmpPrintStatus = IIf(Len(Trim(BordxREc!BordxPrintedStatus)), Trim(BordxREc!BordxPrintedStatus), 0)
        BordxREc.MoveNext
        If tmpPrintStatus = False Then
            CboMonth.Enabled = True
            CboYear.Enabled = True
        '    CboMonth = Format(Date, "mmmm")
        '    CboYear = Year(Date)
        End If
    Loop
    BordxREc.Close
    Set BordxREc = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Function VerifyPrint() As Boolean
    VerifyPrint = False
    If tmpPrintStatus = False Then
        If Trim(CboMonth) = "" Or Trim(CboYear) = "" Then
            MsgBox "Please keyin Month or Year !", vbCritical + vbOKOnly
            CboMonth.SetFocus
            Exit Function
        End If
    End If
    VerifyPrint = True
End Function

Private Function GetFFS(GRPCompany As String) As Boolean
Const METHOD_NAME = "GetFFS"
On Error GoTo debugs

Dim GrpRec As New ADODB.Recordset
Dim GrpSql As String
Dim tmpGrpMcoStatus As String
    GetFFS = False
    tmpGrpMcoStatus = "F"
    If Trim(GRPCompany) <> "" Then
        GrpSql = "Select GRPMCOStatus from GRPCompany where GRPCoName='" & GRPCompany & "'"
        GrpRec.Open GrpSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not GrpRec.EOF
            tmpGrpMcoStatus = IIf(Len(GrpRec!GRPMCOStatus), GrpRec!GRPMCOStatus, "F")
            GrpRec.MoveNext
        Loop
    End If
    GrpRec.Close
    Set GrpRec = Nothing
    If tmpGrpMcoStatus = "T" Then
        GetFFS = True
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Function
Private Sub PrintCoverLetter(tmpcovercompany As String, tmpcoveraadd1 As String, tmpcoveraadd2 As String, tmpcoveraadd3 As String, tmpcoveraadd4 As String, tmpcoveraadd5 As String, tmpcoveraadd6 As String, tmpcoverbordxrefno As String, tmpcoverbordxdate As String, tmpbordxtype As String, tmpbordxcoveramt As String)
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim Add1 As String
Dim Add2 As String
Dim Add3 As String
Dim Add4 As String
Dim Add5 As String
Dim Add6 As String
Dim BordxDate As String
Dim Bordxrefno As String
Dim strsql2 As String
Dim INVPayToset As New ADODB.Recordset

 Set objword = Nothing
 If objword Is Nothing Then
    Set objword = New Word.Application
    objword.Visible = True
    If cboCompany = "Sihat Johor" Then
        Set objdoc = objword.Documents.Add(LocCardPath & "Claim\coverletterSJ.dot")
    Else
        Set objdoc = objword.Documents.Add(LocCardPath & "Claim\coverletter.dot")
    End If
    'objdoc.Bookmarks("companyname").Range.Text = tmpcovercompany
    objdoc.Bookmarks("Add2").Range.Text = Trim(tmpcovercompany) 'MBMGroupCompany
    objdoc.Bookmarks("Add3").Range.Text = Trim(tmpcoveraadd2) 'GRPCoAdd1
    objdoc.Bookmarks("Add4").Range.Text = Trim(tmpcoveraadd3) 'GRPCoAdd2
    objdoc.Bookmarks("Add7").Range.Text = Trim(tmpcoveraadd4) 'GRPCoAdd3
    objdoc.Bookmarks("Add6").Range.Text = Trim(tmpcoveraadd5) 'GRPCoPhone
    objdoc.Bookmarks("Add5").Range.Text = Trim(tmpcoveraadd1) 'GRPCoContact
    objdoc.Bookmarks("Bordxdate").Range.Text = tmpcoverbordxdate 'Bordxdate
    objdoc.Bookmarks("today").Range.Text = Format(Date, "DDmmmmYYYY") 'GETDATE
    
    strsql2 = ""
    tmpbordxtype = ""
    strsql2 = " Select INVPayTo from VIEW_OP_GRPSummary where BordxRefNo = '" & Trim(TxtBordxNo) & "' "
    INVPayToset.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not INVPayToset.EOF
        tmpbordxtype = tmpbordxtype + INVPayToset!INVPayTo + "|"
        INVPayToset.MoveNext
    Loop
    INVPayToset.Close
    Set INVPayToset = Nothing
    
    If Trim(tmpbordxtype) = "C|M|" Then
        objdoc.Bookmarks("Bordxrefno").Range.Text = tmpcoverbordxrefno & "" & " (Cashless/Reimbursement) " & "" & "-" & "" & " RM " & "" & IIf(Len(tmpbordxcoveramt), Format(tmpbordxcoveramt, "#,##0.00"), 0)
    ElseIf Trim(tmpbordxtype) = "C|" Then
        objdoc.Bookmarks("Bordxrefno").Range.Text = tmpcoverbordxrefno & "" & " (Cashless) " & "" & "-" & "" & " RM " & "" & IIf(Len(tmpbordxcoveramt), Format(tmpbordxcoveramt, "#,##0.00"), 0)
    ElseIf Trim(tmpbordxtype) = "M|" Then
        objdoc.Bookmarks("Bordxrefno").Range.Text = tmpcoverbordxrefno & "" & " (Reimbursement) " & "" & "-" & "" & " RM " & "" & IIf(Len(tmpbordxcoveramt), Format(tmpbordxcoveramt, "#,##0.00"), 0)
    End If
 End If
End Sub
Public Sub BYDEPTSJ(objWorkbook As Excel.Workbook)
Dim strSql As String
Dim cmd As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim prm8 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim strsql2 As String
Dim PVCnt As Integer
Dim Rcount As Integer
Dim tmpTotalCases As String
Dim tmpTotalActual As String
Dim tmpTotalHos As Double
Dim tmpHOSPayee As String
Dim tmpSubMBMPaid As Double
Dim RecFound As Boolean
Dim tmpSeqNo As String
Dim totalMBMPaid As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim SubSid As String
Dim Subsidiary As New ADODB.Recordset
Dim Diag As String
Dim sqlbor As String
Dim tmpBordxMonth As String
Dim tmpBordxYear As String
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim SubTotal As Double
Dim SubTotal_App As Double
Dim SubTotal_Copay As Double
Dim IsFirst As Boolean
Dim sValue As String

strSql = ""
                    strSql = " BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd4.ActiveConnection = MediConnCISWS
                    cmd4.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd4.CommandType = adCmdStoredProc
                    cmd4.CommandTimeout = 300
                    Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strSql)
                    cmd4.Parameters.Append Prm7
                    Set prm8 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 6)
                    cmd4.Parameters.Append prm8
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd4.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    RecFound = False
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    
                     strsql2 = ""
                        strsql2 = "SELECT CISDB..View_PrintBordx_Excel.Subsidiary FROM  CISDB..View_PrintBordx_Excel WHERE  BordxRefNo = '" & Trim(TxtBordxNo) & "' and CISDB..View_PrintBordx_Excel.Subsidiary is not null and ltrim(rtrim(CISDB..View_PrintBordx_Excel.Subsidiary)) <> '' GROUP BY CISDB..View_PrintBordx_Excel.Subsidiary"
                        Subsidiary.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                        Do While Not Subsidiary.EOF
                            SubSid = Subsidiary!Subsidiary
                            Subsidiary.MoveNext
                        Loop
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "G") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "G") = TxtBordxNo
                            objWorksheet.Cells(4, "G") = rst2!BordxDate
                        End If
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            objWorksheet.Cells(11 + CDbl(PVCnt), "C") = Format(tmpSeqNo, "000")
                            objWorksheet.Cells(11 + CDbl(PVCnt), "D") = rst2!MBMDepartment
                            objWorksheet.Cells(11 + CDbl(PVCnt), "E") = rst2!TotalCases
                            tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                            objWorksheet.Cells(11 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!TotalActual), "#,##0.00")
                            totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                            objWorksheet.Cells(11 + CDbl(PVCnt), "G") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                        rst2.MoveNext
                    Loop
                        objWorksheet.Cells(12 + CDbl(PVCnt), "D").Font.Bold = True
                        objWorksheet.Cells(12 + CDbl(PVCnt), "D") = "TOTAL"
                        objWorksheet.Cells(12 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(12 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(12 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(12 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                        objWorksheet.Cells(12 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(12 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(12 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(12 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                    rst2.Close
                    Set rst2 = Nothing
                
                Set objWorksheet = objWorkbook.Sheets(1)
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from VIEW_OP_GRPSummaryALL where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
                    tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(12, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(14, "K") = BOR!BordxDate
                    objWorksheet.Cells(12, "B").WrapText = False
                    objWorksheet.Cells(12, "B") = BOR!MBMGroupCompany
                    objWorksheet.Cells(13, "B") = BOR!GRPCoAdd1
                    objWorksheet.Cells(13, "B").WrapText = False
                    objWorksheet.Cells(14, "B") = BOR!GRPCoAdd2
                    objWorksheet.Cells(14, "B").WrapText = False
                    objWorksheet.Cells(15, "B") = BOR!GRPCoAdd3
                    objWorksheet.Cells(15, "B").WrapText = False
                    objWorksheet.Cells(16, "B") = BOR!GRPCoPhone
                    objWorksheet.Cells(16, "B").WrapText = False
                    objWorksheet.Cells(17, "B") = BOR!GRPCoFax
                    objWorksheet.Cells(17, "B").WrapText = False
                    objWorksheet.Cells(18, "C") = Trim(BOR!GRPCoContact)
                    objWorksheet.Cells(18, "C").WrapText = False
                    objWorksheet.Cells(29, "B") = "OUTPATIENT MEDICAL COST FOR DISTRICT" & " " & BOR!MBMGroupCompany
                    tmpcovercompany = IIf(Len(BOR!MBMGroupCompany), BOR!MBMGroupCompany, "")
                    tmpcoveraadd1 = IIf(Len(BOR!GRPCoContact), BOR!GRPCoContact, "NA")
                    tmpcoveraadd2 = IIf(Len(BOR!GRPCoAdd1), BOR!GRPCoAdd1, "NA")
                    tmpcoveraadd3 = IIf(Len(BOR!GRPCoAdd2), BOR!GRPCoAdd2, "NA")
                    tmpcoveraadd4 = IIf(Len(BOR!GRPCoAdd3), BOR!GRPCoAdd3, "NA")
                    tmpcoveraadd5 = IIf(Len(BOR!GRPCoPhone), BOR!GRPCoPhone, "NA")
                    tmpcoveraadd6 = IIf(Len(BOR!GRPCoFax), BOR!GRPCoFax, "NA")
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                objWorksheet.Cells(29, "I") = BOR!TotalVisits
                objWorksheet.Cells(29, "K") = CDbl(BOR!CLMBordxAmt)
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
                
                If RecFound = True Then
                    strSql = ""
                    strSql = " CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 300
                    Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, strSql)
                    cmd2.Parameters.Append Prm3
                    Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 3000, 1)
                    cmd2.Parameters.Append prm4
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd2.Execute
                    PVCnt = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    TMPWS = 1
                    SubTotal = 0
                    SubTotal_App = 0
                    SubTotal_Copay = 0
                    IsFirst = True
                    
                    Do While Not rst2.EOF
                    
                    Diag = ""
                    
                    If IsFirst Then
                        sValue = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                    End If
                    'sValue
                     If sValue = rst2!MBMDepartment Then
                        SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                        SubTotal_Copay = SubTotal_Copay + CDbl(rst2!CLMCoPayColl)
                        'SubTotal_App = SubTotal_App + CDbl(rst2!CLMFinalApprove)
                     Else
                         PVCnt = PVCnt + 2
                         objWorksheet.Cells(9 + CDbl(PVCnt - 1), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(9 + CDbl(PVCnt - 1), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(9 + CDbl(PVCnt - 1), "M").Font.Bold = True
                         objWorksheet.Cells(9 + CDbl(PVCnt - 1), "M") = "Sub Total"
                         
                         objWorksheet.Cells(9 + CDbl(PVCnt - 1), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(9 + CDbl(PVCnt - 1), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(9 + CDbl(PVCnt - 1), "N").Font.Bold = True
                         objWorksheet.Cells(9 + CDbl(PVCnt - 1), "N") = Format(SubTotal_App, "#,##0.00")
                         SubTotal_App = "0.00"
                         sValue = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                     End If
                        
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(3)
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "I") = ": " & rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "I") = ": " & TxtBordxNo
                            objWorksheet.Cells(4, "I") = ": " & rst2!BordxDate
                        End If
                            Rcount = Rcount + 1
                            objWorksheet.Cells(9 + CDbl(PVCnt), "A") = Rcount
                            objWorksheet.Cells(9 + CDbl(PVCnt), "B") = rst2!CasNumber
                            objWorksheet.Cells(9 + CDbl(PVCnt), "C") = IIf(Len(rst2!MBMEmployeeNo), rst2!MBMEmployeeNo, "")
                            objWorksheet.Cells(9 + CDbl(PVCnt), "D") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                            objWorksheet.Cells(9 + CDbl(PVCnt), "E") = rst2!MBMName
                            objWorksheet.Cells(9 + CDbl(PVCnt), "F") = rst2!MBMNumber
                            If rst2!coverid = "00" Then
                                objWorksheet.Cells(9 + CDbl(PVCnt), "G") = rst2!MBMName
                            Else
                                objWorksheet.Cells(9 + CDbl(PVCnt), "G") = rst2!MBMCName
                            End If
                            objWorksheet.Cells(9 + CDbl(PVCnt), "H") = rst2!IC
                            objWorksheet.Cells(9 + CDbl(PVCnt), "I") = rst2!MBMDepartment
                            
                            If Len(rst2!ReasonDesc1) Then
                             Diag = rst2!ReasonDesc1
                            End If
                            
                            If Len(rst2!ReasonDesc2) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc2
                                 Else
                                  Diag = rst2!ReasonDesc2
                                End If
                            End If
                            If Len(rst2!ReasonDesc3) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc3
                                 Else
                                  Diag = rst2!ReasonDesc3
                                End If
                            End If
                            If Len(rst2!CASDiagnosis5) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!CASDiagnosis5
                                Else
                                 Diag = rst2!CASDiagnosis5
                               End If
                            End If
                            objWorksheet.Cells(9 + CDbl(PVCnt), "J") = Diag
                            If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(9 + CDbl(PVCnt), "K") = rst2!NPClinic
                            Else
                                objWorksheet.Cells(9 + CDbl(PVCnt), "K") = rst2!CLNName
                            End If
                            
                            objWorksheet.Cells(9 + CDbl(PVCnt), "L") = IIf(Len(rst2!CLNLocation), rst2!CLNLocation, "")
                            objWorksheet.Cells(9 + CDbl(PVCnt), "M") = Format(rst2!DateVisit, "YYYY/MM/DD")
                            objWorksheet.Cells(9 + CDbl(PVCnt), "N") = Format(rst2!CLMFinalApprove, "#,##0.00")
                            tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMFinalApprove)
                            tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMFinalApprove, "#,##0.00") - Format(rst2!CLMFinalApprove, "#,##0.00"))
                            tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMFinalApprove)
                            SubTotal_App = SubTotal_App + CDbl(rst2!CLMFinalApprove)
                        IsFirst = False
                        rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    RecFound = False
                    
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "M").Font.Bold = True
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "M") = "Sub Total"
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "N").Font.Bold = True
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "N") = Format(SubTotal_App, "#,##0.00")
                         'objWorksheet.Cells(9 + CDbl(PVCnt + 1), "N") = Format(SubTotal, "#,##0.00")
                    PVCnt = PVCnt + 2
                         objWorksheet.Cells(10 + CDbl(PVCnt), "M").Font.Bold = True
                         objWorksheet.Cells(10 + CDbl(PVCnt), "M") = "Total"
                    objWorksheet.Cells(10 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(10 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(10 + CDbl(PVCnt), "N").Font.Bold = True
                    objWorksheet.Cells(10 + CDbl(PVCnt), "N") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(10 + CDbl(PVCnt), "N") = Format(tmpsubAct, "#,##0.00")
                    End If
End Sub

Private Sub PrintSJTemplate()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet

    If cboCompany = "Sihat Johor" And cboBilltype = "Invoice" Then
        If CboReportType.ListIndex = "1" Then
        'Department
                If objExcel Is Nothing Then
                    Set objExcel = CreateObject("Excel.application")
                    objExcel.Interactive = False
                    objExcel.Visible = True
                    objExcel.WindowState = xlMaximized
                    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\SihatJohor_Department.xlt")
                End If
            Call BYDEPTSJ(objWorkbook)
            Call PrintCoverLetter(tmpcovercompany, tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
              objExcel.DisplayAlerts = True
              objExcel.Interactive = True
              objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=1, OpenAfterPublish:=True
        ElseIf CboReportType.ListIndex = "5" Then
        'Subsidiry
            MsgBox "Subsidiry Report Type Still Under Progress..!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            MsgBox "Report Type Still Under Progress..!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    Else
        MsgBox "Report Type Still Under Progress..!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Function GetCompanyCodeByBordx() As String
    Dim rsCompany As New ADODB.Recordset
    Dim cmdCompany As New ADODB.Command
    
    cmdCompany.ActiveConnection = MediConnCISWS
    cmdCompany.CommandText = "spGetCompanyByBordx"
    cmdCompany.CommandType = adCmdStoredProc
    cmdCompany.CommandTimeout = 300
    
    Dim prmBordx As New ADODB.Parameter
    Set prmBordx = cmdCompany.CreateParameter("Bordx", adVarChar, adParamInput, 10, Trim(TxtBordxNo.Text))
    cmdCompany.Parameters.Append prmBordx

    Set rsCompany = cmdCompany.Execute
    
    If rsCompany.BOF = True And rsCompany.EOF = True Then
        GetCompanyCodeByBordx = ""
    Else
        GetCompanyCodeByBordx = rsCompany!GRPCoCode
    End If
    rsCompany.Close
End Function


Private Function GetDataFromCISWS(strSql As String, strCase As String, strSP) As Recordset

    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command

    cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = strSP '"CLM_PayAdviceCLNBSN"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300

    Dim prmSql As New ADODB.Parameter
    Set prmSql = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strSql)
    'Set prmSql = cmdPayAdvise.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, " BordxRefNo = '" & Trim(TxtBordxNo) & "'")
    cmd.Parameters.Append prmSql
    
    Dim prmCase As New ADODB.Parameter
    Set prmCase = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strCase)
    cmd.Parameters.Append prmCase
    
    Set rs = cmd.Execute

    Set GetDataFromCISWS = rs

End Function


Private Sub PrintTemplate()
Dim StrSql1, strsql2 As String
Dim Subsidiary As New ADODB.Recordset
Dim strsqlPlutoBor As String
Dim PlutoBor As New ADODB.Recordset
Dim sqlinsert As String
Dim sqlbor As String
Dim tmpbordxdate As String
Dim tmpDateVisit As String
Dim TempMBMName, TempMBMDept, tempPath As String
Dim PLUTOStatus As Boolean
Dim BOR As New ADODB.Recordset
Dim TempRemarks As String
Dim tempGrpCompany As String
Dim TempCASDiagnosis5, TempPatientName, TempMBMCName, TempCLNName As String
Dim tmpSeqNo, PVCnt As Integer
Dim tmpTotalBordxAmt As String
Dim tmpTotalActual, tmptotalcopay, tmpTotalinvamt, tmpTotalVisits, tmpTotalMC, tmpHosName As String
Dim tmpTotalMBMDue, tmpTotalHos As Double
Dim TMPWS As Integer
Dim tmpSubApp, tmpsubAct, tmpsubcopay, tmpsubinvamt As Double
Dim strsqlHOS, tmpHOSPayee As String
Dim HOS As New ADODB.Recordset
Dim strsqlCOV, tmpPatientName, tmpPatientRel As String
Dim RstCov As New ADODB.Recordset
Dim tmpSubMBMPaid, totalMBMPaid, PVReimCnt As Double
Dim PLN As New ADODB.Recordset
Dim strsqlPLN, PlnDesc As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim strSql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim cmd3 As New ADODB.Command
Dim Prm5 As New ADODB.Parameter
Dim prm6 As New ADODB.Parameter
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim prm8 As New ADODB.Parameter
Dim cmd5 As New ADODB.Command
Dim Prm9 As New ADODB.Parameter
Dim prm10 As New ADODB.Parameter
Dim tmpTotalCases, lblDate, sValue As String
Dim RecFound, IsFirst As Boolean
Dim SubTotal, SubTotalcopay, SubTotalinvamt As Double
Dim Rcount As Integer
Dim Diag, SubSid, tmpBordxMonth, tmpBordxYear As String

SubSid = ""
tmpcovercompany = ""
tmpcoveraadd1 = ""
tmpcoveraadd2 = ""
tmpcoveraadd3 = ""
tmpcoveraadd4 = ""
tmpcoveraadd5 = ""
tmpcoveraadd6 = ""
tmpcoverbordxrefno = ""
tmpcoverbordxdate = ""
tmpbordxtype = ""
tmpbordxcoveramt = ""
Diag = ""
IsFirst = True
    
Screen.MousePointer = vbHourglass

    If TxtBordxNo = "" Then
    
        MsgBox "Please Select Bordx", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            compCodeByBordx = GetCompanyCodeByBordx()
            If cboCompany = "MediExpress" Then
                If CboReportType.ListIndex = "0" Then
                    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PayAdviceCLNBordx.xlt")
                'ElseIf CboReportType.ListIndex = "1" Then
                 '   Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PayAdviceCLNBordxHCDept1.xlt")
                'ElseIf CboReportType.ListIndex = "3" Then
                 '   Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BordxByCostCenter.xlt")
                'ElseIf CboReportType.ListIndex = "4" Then
                 '   Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BordxByCostCenter-Branch.xlt")
                'ElseIf CboReportType.ListIndex = "5" Then
                 '   Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\New_PayAdviceCLNBordxBYSUB.xlt")
                Else
                    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PayAdviceCLNBordxDept.xlt")
                    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PayAdviceCLNBordxRel.xlt")
                End If
            Else
                If CboReportType.ListIndex = "0" Then
                    If Mid(TxtBordxNo, 1, 3) = "KCN" Then
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\CreditNote_New_PayAdviceCLNBordx_20.xlt")
                    ElseIf cboBilltype.ListIndex = "0" Then
                        'Tax Invoice
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PayAdviceCLNBordx_TAX.xlt")
                    ElseIf cboBilltype.ListIndex = "1" Then
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PayAdviceCLNBordx_20.xlt")
                    End If
                ElseIf CboReportType.ListIndex = "1" Then
                     If Mid(TxtBordxNo, 1, 3) = "KCN" Then
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\CreditNote_New_PayAdviceCLNBordxHCDept_20.xlt")
                     ElseIf cboBilltype.ListIndex = "0" Then
                        'Tax Invoice
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\New_PayAdviceCLNBordxHCDept_TAX.xlt")
                     ElseIf cboBilltype.ListIndex = "1" Then
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\New_PayAdviceCLNBordxHCDept_20.xlt")
                     End If
                     
                ElseIf CboReportType.ListIndex = "2" Then
                     If Mid(TxtBordxNo, 1, 3) = "KCN" Then
                     Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\CreditNote_New_PayAdviceCLNBordxRel_2020.xlt")
                     ElseIf cboBilltype.ListIndex = "0" Then
                        'Tax Invoice
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PayAdviceCLNBordxRel_TAX_2020.xlt")
                     ElseIf cboBilltype.ListIndex = "1" Then
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PayAdviceCLNBordxRel_2020.xlt")
                     End If
                     
                ElseIf CboReportType.ListIndex = "3" Then
                     If Mid(TxtBordxNo, 1, 3) = "KCN" Then
                     Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\CreditNote_New_BordxByCostCenter_20.xlt")
                     ElseIf cboBilltype.ListIndex = "0" Then
                        'Tax Invoice
                        If compCodeByBordx = "P257" Then 'Customised template for PureCircle Client
                            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BordxByCostCenter_TAX_PureCircle.xlt")
                        Else
                            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BordxByCostCenter_TAX.xlt")
                        End If
                     ElseIf cboBilltype.ListIndex = "1" Then 'Customised template for PureCircle Client
                        If compCodeByBordx = "P257" Then
                            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BordxByCostCenter_20_PureCircle.xlt")
                        Else
                            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BordxByCostCenter_20.xlt")
                        End If
                     End If
                     
                ElseIf CboReportType.ListIndex = "4" Then
                     If Mid(TxtBordxNo, 1, 3) = "KCN" Then
                     Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\CreditNote_New_BordxByCostCenter-Branch_20.xlt")
                     ElseIf cboBilltype.ListIndex = "0" Then
                        'Tax Invoice
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BordxByCostCenter-Branch_TAX.xlt")
                     ElseIf cboBilltype.ListIndex = "1" Then
                     Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BordxByCostCenter-Branch_20.xlt")
                     End If
                     
               Else
                     If Mid(TxtBordxNo, 1, 3) = "KCN" Then
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\CreditNote_New_PayAdviceCLNBordxBYSUB.xlt")
                     ElseIf cboBilltype.ListIndex = "0" Then
                        'Tax Invoice
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\New_PayAdviceCLNBordxBYSUB_TAX.xlt")
                     ElseIf cboBilltype.ListIndex = "1" Then
                        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\New_PayAdviceCLNBordxBYSUB.xlt")
                     End If
                     
               End If
            End If
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSubMBMPaid = 0
        RecFound = False
        tmpSeqNo = 0
        totalMBMPaid = 0
        PVReimCnt = 0
        tmpPatientRel = ""
        Diag = ""
            If CboReportType.ListIndex = "0" Then
                    strSql = ""
                    Set rst2 = GetDataFromCISWS("  BordxRefNo = '" & Trim(TxtBordxNo) & "'", 4, "CLM_PayAdviceCLNBSN")
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    RecFound = False
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                        strsql2 = ""
                        strsql2 = "SELECT CISDB..View_PrintBordx_Excel.Subsidiary FROM  CISDB..View_PrintBordx_Excel WHERE  BordxRefNo = '" & Trim(TxtBordxNo) & "' and CISDB..View_PrintBordx_Excel.Subsidiary is not null and ltrim(rtrim(CISDB..View_PrintBordx_Excel.Subsidiary)) <> '' GROUP BY CISDB..View_PrintBordx_Excel.Subsidiary"
                        Subsidiary.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                        Do While Not Subsidiary.EOF
                            SubSid = Subsidiary!Subsidiary
                            Subsidiary.MoveNext
                        Loop
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        
                    Set objWorksheet = objWorkbook.Sheets(1)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from VIEW_OP_GRPSummaryALL where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
        
                   tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(12, "K") = BOR!BordxDate
                    objWorksheet.Cells(16, "C") = Trim(BOR!GRPCoContact)
                    objWorksheet.Cells(16, "C").WrapText = False
                    objWorksheet.Cells(10, "B").WrapText = False
                    objWorksheet.Cells(10, "B") = BOR!MBMGroupCompany
                    objWorksheet.Cells(11, "B") = BOR!GRPCoAdd1
                    objWorksheet.Cells(11, "B").WrapText = False
                    objWorksheet.Cells(12, "B") = BOR!GRPCoAdd2
                    objWorksheet.Cells(12, "B").WrapText = False
                    objWorksheet.Cells(13, "B") = BOR!GRPCoAdd3
                    objWorksheet.Cells(13, "B").WrapText = False
                    objWorksheet.Cells(14, "B") = BOR!GRPCoPhone
                    objWorksheet.Cells(14, "B").WrapText = False
                    objWorksheet.Cells(15, "B") = BOR!GRPCoFax
                    objWorksheet.Cells(15, "B").WrapText = False
                    objWorksheet.Cells(27, "B") = "OUTPATIENT MEDICAL COST FOR" & " " & BOR!MBMGroupCompany
                    objWorksheet.Cells(27, "I") = BOR!TotalVisits
                    objWorksheet.Cells(27, "K") = CDbl(BOR!CLMBordxAmt)
                    
                    tmpcovercompany = IIf(Len(BOR!MBMGroupCompany), BOR!MBMGroupCompany, "")
                    tmpcoveraadd1 = IIf(Len(BOR!GRPCoContact), BOR!GRPCoContact, "NA")
                    tmpcoveraadd2 = IIf(Len(BOR!GRPCoAdd1), BOR!GRPCoAdd1, "NA")
                    tmpcoveraadd3 = IIf(Len(BOR!GRPCoAdd2), BOR!GRPCoAdd2, "NA")
                    tmpcoveraadd4 = IIf(Len(BOR!GRPCoAdd3), BOR!GRPCoAdd3, "NA")
                    tmpcoveraadd5 = IIf(Len(BOR!GRPCoPhone), BOR!GRPCoPhone, "NA")
                    tmpcoveraadd6 = IIf(Len(BOR!GRPCoFax), BOR!GRPCoFax, "NA")
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                    BOR.MoveNext
                    Loop
                    BOR.Close
                    Set BOR = Nothing
                        Set objWorksheet = objWorkbook.Sheets(2)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                                               
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "H") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "H") = TxtBordxNo
                            objWorksheet.Cells(4, "H") = rst2!BordxDate
                            objWorksheet.Cells(5, "H") = IIf(Len(SubSid), SubSid, "")
                        End If
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            objWorksheet.Cells(16 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                            objWorksheet.Cells(16 + CDbl(PVCnt), "B") = rst2!MBMDepartment
                            objWorksheet.Cells(16 + CDbl(PVCnt), "E") = rst2!TotalCases
                            tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                            objWorksheet.Cells(16 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!TotalActual), "#,##0.00")
                            totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                            objWorksheet.Cells(16 + CDbl(PVCnt), "G") = Format(rst2!PayabletoHOS, "#,##0.00")
                            objWorksheet.Cells(16 + CDbl(PVCnt), "H") = CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS)
                            
                            tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                            tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
                        rst2.MoveNext
                    Loop
                        
                        objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(18 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(18 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                        objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(18 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(18 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                        objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
                        objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(totalMBMPaid, "#,##0.00")
                        objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
                        objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(tmpTotalHos, "#,##0.00")
                    rst2.Close
                    Set rst2 = Nothing
                    
                If RecFound = True Then
                    strSql = ""
                    strSql = " CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 300
                    Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, strSql)
                    cmd2.Parameters.Append Prm3
                    Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 3000, 1)
                    cmd2.Parameters.Append prm4
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd2.Execute
                    PVCnt = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    TMPWS = 1
                    SubTotal = 0
                    
                    Do While Not rst2.EOF
                    
                    If IsFirst Then
                    sValue = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                    End If
                    
                     If sValue = rst2!MBMDepartment Then
                        SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                        Else
                         PVCnt = PVCnt + 2
                         objWorksheet.Cells(14 + CDbl(PVCnt - 1), "U").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(14 + CDbl(PVCnt - 1), "U").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(14 + CDbl(PVCnt - 1), "U").Font.Bold = True
                         objWorksheet.Cells(14 + CDbl(PVCnt - 1), "U") = "Sub Total"
                         objWorksheet.Cells(14 + CDbl(PVCnt - 1), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(14 + CDbl(PVCnt - 1), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(14 + CDbl(PVCnt - 1), "V").Font.Bold = True
                         objWorksheet.Cells(14 + CDbl(PVCnt - 1), "V") = Format(SubTotal, "#,##0.00")
                         sValue = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                         SubTotal = 0
                         SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                        End If
                        
                        Diag = ""
                        PVCnt = PVCnt + 1
                        PlnDesc = ""
                        strsqlPLN = "Select PLNCode ,PLNDescription from PLN where PLNCode = '" & rst2!HISPLN & "'"
                        PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                        If PLN.RecordCount > 0 Then
                            PlnDesc = Mid(PLN!PLNDescription, InStr(1, PLN!PLNDescription, "-") + 1)
                        End If
                        PLN.Close
                        Set PLN = Nothing
                        Set objWorksheet = objWorkbook.Sheets(3)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "L") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "L") = TxtBordxNo
                            objWorksheet.Cells(4, "L") = rst2!BordxDate
                            objWorksheet.Cells(6, "L") = IIf(Len(SubSid), SubSid, "")
                        End If
                            Rcount = Rcount + 1
                            objWorksheet.Cells(14 + CDbl(PVCnt), "A") = Rcount
                            objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CasNumber
                            objWorksheet.Cells(14 + CDbl(PVCnt), "C") = Right(PlnDesc, 2)
                            If rst2!coverid = "00" Then
                                objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "E"
                            Else
                                If rst2!MBMCRELCode = "H" Then
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "S"
                                ElseIf rst2!MBMCRELCode = "W" Then
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "S"
                                ElseIf rst2!MBMCRELCode = "S" Then
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "C"
                                ElseIf rst2!MBMCRELCode = "D" Then
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "C"
                                End If
                            End If
                            
                            objWorksheet.Cells(14 + CDbl(PVCnt), "F") = rst2!MBMDepartment
                            objWorksheet.Cells(14 + CDbl(PVCnt), "G") = rst2!MBMName
                            If rst2!coverid = "00" Then
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I") = rst2!MBMName
                            Else
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I") = rst2!MBMCName
                            End If
                            If Len(rst2!ReasonDesc1) Then
                             Diag = rst2!ReasonDesc1
                            End If
                            If Len(rst2!ReasonDesc2) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc2
                                 Else
                                  Diag = rst2!ReasonDesc2
                                End If
                            End If
                            If Len(rst2!ReasonDesc3) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc3
                                 Else
                                  Diag = rst2!ReasonDesc3
                                End If
                            End If
                            If Len(rst2!CASDiagnosis5) Then
                                If Len(Diag) Then
                                Diag = Diag & " \ " & rst2!CASDiagnosis5
                                Else
                                 Diag = rst2!CASDiagnosis5
                               End If
                            End If
                            
                            objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Diag
                            If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M") = rst2!NPClinic
                            Else
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M") = rst2!CLNName
                            End If
                            objWorksheet.Cells(14 + CDbl(PVCnt), "P") = rst2!CLNLocation
                            objWorksheet.Cells(14 + CDbl(PVCnt), "Q") = rst2!MBMNumber
                            objWorksheet.Cells(14 + CDbl(PVCnt), "R") = Format(rst2!DateVisit, "YYYY/MM/DD")
                            objWorksheet.Cells(14 + CDbl(PVCnt), "S") = IIf(Len(rst2!MBMEmployeeNo), rst2!MBMEmployeeNo, "")
                            objWorksheet.Cells(14 + CDbl(PVCnt), "T") = rst2!IC
                            objWorksheet.Cells(14 + CDbl(PVCnt), "U") = Format(rst2!CLMTotalActual, "#,##0.00")
                            tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                            tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMTotalActual, "#,##0.00"))
                            objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(rst2!CLMTotalActual, "#,##0.00")
                            tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                       IsFirst = False
                       rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    RecFound = False
                    
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "U").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "U").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "U").Font.Bold = True
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "U") = "Sub Total"
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Font.Bold = True
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V") = Format(SubTotal, "#,##0.00")
                    PVCnt = PVCnt + 2
                    objWorksheet.Cells(14 + CDbl(PVCnt), "U").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "U").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "U").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "U") = "Total"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(tmpSubApp, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(tmpSubApp, "#,##0.00")
                    Call PrintCoverLetter(tmpcovercompany, tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
            End If
        
        ElseIf CboReportType.ListIndex = "1" Then
           Call BYDEPT(objWorkbook)
           Call PrintCoverLetter(tmpcovercompany, tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
        ElseIf CboReportType.ListIndex = "3" Then
        
        Set rst2 = GetDataFromCISWS(" BordxRefNo = '" & Trim(TxtBordxNo) & "'", 10, "CLM_PayAdviceCLNBSN")
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    RecFound = False
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                     strsql2 = ""
                        strsql2 = "SELECT CISDB..View_PrintBordx_Excel.Subsidiary FROM  CISDB..View_PrintBordx_Excel WHERE  BordxRefNo = '" & Trim(TxtBordxNo) & "' and CISDB..View_PrintBordx_Excel.Subsidiary is not null and ltrim(rtrim(CISDB..View_PrintBordx_Excel.Subsidiary)) <> '' GROUP BY CISDB..View_PrintBordx_Excel.Subsidiary"
                        Subsidiary.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                        Do While Not Subsidiary.EOF
                            SubSid = Subsidiary!Subsidiary
                            Subsidiary.MoveNext
                        Loop
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(1)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from VIEW_OP_GRPSummaryALL where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
                   tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(12, "K") = BOR!BordxDate
                    objWorksheet.Cells(16, "C") = Trim(BOR!GRPCoContact)
                    objWorksheet.Cells(16, "C").WrapText = False
                    objWorksheet.Cells(10, "B").WrapText = False
                    objWorksheet.Cells(10, "B") = IIf(Len(SubSid) > 0, SubSid, BOR!MBMGroupCompany)
                    objWorksheet.Cells(11, "B") = BOR!GRPCoAdd1
                    objWorksheet.Cells(11, "B").WrapText = False
                    objWorksheet.Cells(12, "B") = BOR!GRPCoAdd2
                    objWorksheet.Cells(12, "B").WrapText = False
                    objWorksheet.Cells(13, "B") = BOR!GRPCoAdd3
                    objWorksheet.Cells(13, "B").WrapText = False
                    objWorksheet.Cells(14, "B") = BOR!GRPCoPhone
                    objWorksheet.Cells(14, "B").WrapText = False
                    objWorksheet.Cells(15, "B") = BOR!GRPCoFax
                    objWorksheet.Cells(15, "B").WrapText = False
                    objWorksheet.Cells(27, "B") = "OUTPATIENT MEDICAL COST FOR" & " " & BOR!MBMGroupCompany
                    objWorksheet.Cells(27, "I") = BOR!TotalVisits
                    objWorksheet.Cells(27, "K") = CDbl(BOR!CLMBordxAmt)
                    tmpcovercompany = IIf(Len(BOR!MBMGroupCompany), BOR!MBMGroupCompany, "")
                    tmpcoveraadd1 = IIf(Len(BOR!GRPCoContact), BOR!GRPCoContact, "NA")
                    tmpcoveraadd2 = IIf(Len(BOR!GRPCoAdd1), BOR!GRPCoAdd1, "NA")
                    tmpcoveraadd3 = IIf(Len(BOR!GRPCoAdd2), BOR!GRPCoAdd2, "NA")
                    tmpcoveraadd4 = IIf(Len(BOR!GRPCoAdd3), BOR!GRPCoAdd3, "NA")
                    tmpcoveraadd5 = IIf(Len(BOR!GRPCoPhone), BOR!GRPCoPhone, "NA")
                    tmpcoveraadd6 = IIf(Len(BOR!GRPCoFax), BOR!GRPCoFax, "NA")
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
                        
                  Set objWorksheet = objWorkbook.Sheets(2)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "K") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "K") = TxtBordxNo
                            objWorksheet.Cells(4, "K") = rst2!BordxDate
                            objWorksheet.Cells(5, "K") = IIf(Len(SubSid), SubSid, "")
                        End If
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                            If compCodeByBordx = "P257" Then
                                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = rst2!MBMDivision
                            End If
                            objWorksheet.Cells(7 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "C", "B")) = rst2!MBMDepartment
                            objWorksheet.Cells(7 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "D", "C")) = rst2!MBMCostCenter
                            objWorksheet.Cells(7 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "E", "D")) = rst2!TotalCases
                            objWorksheet.Cells(7 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "F", "E")) = rst2!TotalActual
                            
                            objWorksheet.Cells(7 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "G", "F")) = rst2!CLMCoPayCLNColl
                            objWorksheet.Cells(7 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "H", "G")) = rst2!CLMBordxAmt
                            
                            objWorksheet.Cells(7 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "I", "H")) = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                            tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!TotalActual), "#,##0.00")
                            
                            tmptotalcopay = CDbl(tmptotalcopay) + Format(CDbl(rst2!CLMCoPayCLNColl), "#,##0.00")
                            tmpTotalinvamt = CDbl(tmpTotalinvamt) + Format(CDbl(rst2!CLMBordxAmt), "#,##0.00")
                            
                        rst2.MoveNext
                    Loop
                        objWorksheet.Range(IIf(compCodeByBordx = "P257", "D", "C") + CStr(8 + CDbl(PVCnt)) + ":" + IIf(compCodeByBordx = "P257", "H", "G") + CStr(8 + CDbl(PVCnt))).Font.Bold = True
                        objWorksheet.Range(IIf(compCodeByBordx = "P257", "D", "C") + CStr(8 + CDbl(PVCnt)) + ":" + IIf(compCodeByBordx = "P257", "H", "G") + CStr(8 + CDbl(PVCnt))).Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Range(IIf(compCodeByBordx = "P257", "D", "C") + CStr(8 + CDbl(PVCnt)) + ":" + IIf(compCodeByBordx = "P257", "H", "G") + CStr(8 + CDbl(PVCnt))).Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "D", "C")) = "Total"
                        objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "E", "D")) = Format(tmpTotalCases, "#,##0.00")
                        objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "F", "E")) = Format(tmpTotalActual, "#,##0.00")
                        objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "G", "F")) = Format(tmptotalcopay, "#,##0.00")
                        objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "H", "G")) = Format(tmpTotalinvamt, "#,##0.00")
                    rst2.Close
                    Set rst2 = Nothing
                    
                   
                If RecFound = True Then
                    Set rst2 = GetDataFromCISWS(" CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'", 7, "CLM_PayAdviceCLNBSN")
                    PVCnt = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    TMPWS = 1
                    SubTotal = 0
                    Dim TmpSval As String
                    Do While Not rst2.EOF
                    If IsFirst Then
                    If IsNull(rst2!MBMCostCenter) Or rst2!MBMCostCenter = "Null" Then
                        sValue = "Nul"
                    Else
                        sValue = rst2!MBMCostCenter
                    End If
                        sValue = IIf(IsNull(rst2!MBMCostCenter), "Nul", rst2!MBMCostCenter)
                    End If
                    
                    If IsNull(rst2!MBMCostCenter) Or rst2!MBMCostCenter = "Null" Or Trim(rst2!MBMCostCenter) = Trim("") Then
                        TmpSval = "Nul"
                    Else
                        TmpSval = rst2!MBMCostCenter
                    End If
                    
                     If sValue = TmpSval Then
                        SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                        SubTotalcopay = SubTotalcopay + CDbl(rst2!CLMCoPayCLNColl)
                        SubTotalinvamt = SubTotalinvamt + CDbl(rst2!CLMBordxAmt)
                        Else
                         PVCnt = PVCnt + 2
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), IIf(compCodeByBordx = "P257", "P", "O")) = "Sub Total"
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), IIf(compCodeByBordx = "P257", "Q", "P")) = Format(SubTotal, "#,##0.00")
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), IIf(compCodeByBordx = "P257", "R", "Q")) = Format(SubTotalcopay, "#,##0.00")
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), IIf(compCodeByBordx = "P257", "S", "R")) = Format(SubTotalinvamt, "#,##0.00")
                         
                         objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(8 + CDbl(PVCnt - 1)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(8 + CDbl(PVCnt - 1))).Font.Bold = True
                         objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(8 + CDbl(PVCnt - 1)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(8 + CDbl(PVCnt - 1))).Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(8 + CDbl(PVCnt - 1)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(8 + CDbl(PVCnt - 1))).Borders(xlEdgeBottom).LineStyle = xlDouble
                         
                         
                         sValue = rst2!MBMCostCenter
                         SubTotal = 0
                         SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                         SubTotalcopay = 0
                         SubTotalcopay = SubTotalcopay + CDbl(rst2!CLMCoPayCLNColl)
                         SubTotalinvamt = 0
                         SubTotalinvamt = SubTotalinvamt + CDbl(rst2!CLMBordxAmt)
                        End If
                        
                        Diag = ""
                        PVCnt = PVCnt + 1
                        
                        Set objWorksheet = objWorkbook.Sheets(3)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                        
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "M") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "M") = TxtBordxNo
                            objWorksheet.Cells(4, "M") = rst2!BordxDate
                            objWorksheet.Cells(5, "M") = IIf(Len(SubSid), SubSid, "")
                        End If
                            Rcount = Rcount + 1
                            If Len(rst2!ReasonDesc1) Then
                             Diag = rst2!ReasonDesc1
                            End If
                            If Len(rst2!ReasonDesc2) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc2
                                 Else
                                  Diag = rst2!ReasonDesc2
                                End If
                            End If
                            If Len(rst2!ReasonDesc3) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc3
                                 Else
                                  Diag = rst2!ReasonDesc3
                                End If
                            End If
                             If Len(rst2!ReasonDesc4) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc4
                                 Else
                                  Diag = rst2!ReasonDesc4
                                End If
                            End If
                            If Len(rst2!CASDiagnosis5) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!CASDiagnosis5
                                Else
                                 Diag = rst2!CASDiagnosis5
                             End If
                           End If
                            objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Rcount
                            objWorksheet.Cells(8 + CDbl(PVCnt), "B") = rst2!CasNumber
                            objWorksheet.Cells(8 + CDbl(PVCnt), "C") = rst2!MBMNumber
                            
                            If compCodeByBordx = "P257" Then
                                objWorksheet.Cells(8 + CDbl(PVCnt), "D") = rst2!MBMDivision
                            End If
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "E", "D")) = rst2!MBMDepartment
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "F", "E")) = rst2!MBMCostCenter
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "G", "F")) = rst2!MBMEmployeeNo
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "H", "G")) = rst2!MBMName
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "I", "H")) = rst2!IC
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "J", "I")) = rst2!PatientName
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "K", "J")) = rst2!DateVisit
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "M", "K")) = Diag
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "N", "M")) = rst2!CLNType
                            If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "O", "N")) = rst2!NPClinic
                            Else
                                objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "O", "N")) = rst2!CLNName
                            End If
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "P", "O")) = rst2!CLNLocation
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "Q", "P")) = rst2!CLMTotalActual
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "R", "Q")) = rst2!CLMCoPayCLNColl
                            objWorksheet.Cells(8 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "S", "R")) = rst2!CLMBordxAmt
                            tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                            
                            tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                            tmpsubcopay = CDbl(tmpsubcopay) + CDbl(rst2!CLMCoPayCLNColl)
                            tmpsubinvamt = CDbl(tmpsubinvamt) + CDbl(rst2!CLMBordxAmt)
                       IsFirst = False
                       rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    RecFound = False
                    
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), IIf(compCodeByBordx = "P257", "P", "O")) = "Sub Total"
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), IIf(compCodeByBordx = "P257", "Q", "P")) = Format(SubTotal, "#,##0.00")
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), IIf(compCodeByBordx = "P257", "R", "Q")) = Format(SubTotalcopay, "#,##0.00")
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), IIf(compCodeByBordx = "P257", "S", "R")) = Format(SubTotalinvamt, "#,##0.00")
                    
                    objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(8 + CDbl(PVCnt + 1)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(8 + CDbl(PVCnt + 1))).Font.Bold = True
                    objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(8 + CDbl(PVCnt + 1)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(8 + CDbl(PVCnt + 1))).Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(8 + CDbl(PVCnt + 1)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(8 + CDbl(PVCnt + 1))).Borders(xlEdgeBottom).LineStyle = xlDouble
                    
                    PVCnt = PVCnt + 1
                    
                    objWorksheet.Cells(11 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "P", "O")) = "Total"
                    objWorksheet.Cells(11 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "Q", "P")) = Format(tmpSubApp, "#,##0.00")
                    objWorksheet.Cells(11 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "R", "Q")) = Format(tmpsubcopay, "#,##0.00")
                    objWorksheet.Cells(11 + CDbl(PVCnt), IIf(compCodeByBordx = "P257", "S", "R")) = Format(tmpsubinvamt, "#,##0.00")
                    
                    objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(11 + CDbl(PVCnt)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(11 + CDbl(PVCnt))).Font.Bold = True
                    objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(11 + CDbl(PVCnt)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(11 + CDbl(PVCnt))).Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Range(IIf(compCodeByBordx = "P257", "P", "O") + CStr(11 + CDbl(PVCnt)) + ":" + IIf(compCodeByBordx = "P257", "S", "R") + CStr(11 + CDbl(PVCnt))).Borders(xlEdgeBottom).LineStyle = xlDouble
            
            Call PrintCoverLetter(IIf(Len(SubSid), SubSid, tmpcovercompany), tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
            End If
        
        ElseIf CboReportType.ListIndex = "4" Then
                    strSql = ""
                    Set rst2 = GetDataFromCISWS(" BordxRefNo = '" & Trim(TxtBordxNo) & "'", 8, "CLM_PayAdviceCLNBSN")
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    RecFound = False
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    strsql2 = ""
                    
                        strsql2 = "SELECT CISDB..View_PrintBordx_Excel.Subsidiary FROM  CISDB..View_PrintBordx_Excel WHERE  BordxRefNo = '" & Trim(TxtBordxNo) & "' and CISDB..View_PrintBordx_Excel.Subsidiary is not null and ltrim(rtrim(CISDB..View_PrintBordx_Excel.Subsidiary)) <> '' GROUP BY CISDB..View_PrintBordx_Excel.Subsidiary"
                        Subsidiary.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                        Do While Not Subsidiary.EOF
                            SubSid = Subsidiary!Subsidiary
                            Subsidiary.MoveNext
                        Loop
                    
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        
                        Set objWorksheet = objWorkbook.Sheets(1)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from VIEW_OP_GRPSummaryALL where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
        
                   tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(12, "K") = BOR!BordxDate
                    objWorksheet.Cells(16, "C") = Trim(BOR!GRPCoContact)
                    objWorksheet.Cells(16, "C").WrapText = False
                    objWorksheet.Cells(10, "B").WrapText = False
                    objWorksheet.Cells(10, "B") = BOR!MBMGroupCompany
                    objWorksheet.Cells(11, "B") = BOR!GRPCoAdd1
                    objWorksheet.Cells(11, "B").WrapText = False
                    objWorksheet.Cells(12, "B") = BOR!GRPCoAdd2
                    objWorksheet.Cells(12, "B").WrapText = False
                    objWorksheet.Cells(13, "B") = BOR!GRPCoAdd3
                    objWorksheet.Cells(13, "B").WrapText = False
                    objWorksheet.Cells(14, "B") = BOR!GRPCoPhone
                    objWorksheet.Cells(14, "B").WrapText = False
                    objWorksheet.Cells(15, "B") = BOR!GRPCoFax
                    objWorksheet.Cells(15, "B").WrapText = False
                    objWorksheet.Cells(27, "B") = "OUTPATIENT MEDICAL COST FOR" & " " & BOR!MBMGroupCompany
                    objWorksheet.Cells(27, "I") = BOR!TotalVisits
                    objWorksheet.Cells(27, "K") = CDbl(BOR!CLMBordxAmt)
                
                    tmpcovercompany = IIf(Len(BOR!MBMGroupCompany), BOR!MBMGroupCompany, "")
                    tmpcoveraadd1 = IIf(Len(BOR!GRPCoContact), BOR!GRPCoContact, "NA")
                    tmpcoveraadd2 = IIf(Len(BOR!GRPCoAdd1), BOR!GRPCoAdd1, "NA")
                    tmpcoveraadd3 = IIf(Len(BOR!GRPCoAdd2), BOR!GRPCoAdd2, "NA")
                    tmpcoveraadd4 = IIf(Len(BOR!GRPCoAdd3), BOR!GRPCoAdd3, "NA")
                    tmpcoveraadd5 = IIf(Len(BOR!GRPCoPhone), BOR!GRPCoPhone, "NA")
                    tmpcoveraadd6 = IIf(Len(BOR!GRPCoFax), BOR!GRPCoFax, "NA")
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                    
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
                        Set objWorksheet = objWorkbook.Sheets(2)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "K") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "K") = TxtBordxNo
                            objWorksheet.Cells(4, "K") = rst2!BordxDate
                            objWorksheet.Cells(5, "K") = IIf(Len(SubSid), SubSid, "")
                        End If
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "B") = rst2!MBMBranch
                            objWorksheet.Cells(7 + CDbl(PVCnt), "C") = IIf(Len(rst2!MBMCostCenter), rst2!MBMCostCenter, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "D") = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "E") = IIf(Len(rst2!MBMCostCenter), rst2!MBMCostCenter, "") + IIf(Len(CStr(rst2!MBMDepartment)), CStr(rst2!MBMDepartment), "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "F") = rst2!TotalActual
                            tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!TotalActual), "#,##0.00")
                        rst2.MoveNext
                    Loop
                        
                        objWorksheet.Cells(8 + CDbl(PVCnt), "D").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(8 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(8 + CDbl(PVCnt), "D").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "D") = "Total"
                        objWorksheet.Cells(8 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(8 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(8 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                        objWorksheet.Cells(8 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(8 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(8 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                    rst2.Close
                    Set rst2 = Nothing
                    
                If RecFound = True Then
                    strSql = ""
                    Set rst2 = GetDataFromCISWS(" CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'", 7, "CLM_PayAdviceCLNBSN")
                            
                    PVCnt = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    TMPWS = 1
                    SubTotal = 0
                    'Dim TmpSval As String
                    Do While Not rst2.EOF
                    If IsFirst Then
                    If IsNull(rst2!MBMCostCenter) Or rst2!MBMCostCenter = "Null" Then
                        sValue = "Nul"
                    Else
                        sValue = rst2!MBMCostCenter
                    End If
                    sValue = IIf(IsNull(rst2!MBMCostCenter), "Nul", rst2!MBMCostCenter)
                    End If
                    
                    If IsNull(rst2!MBMCostCenter) Or rst2!MBMCostCenter = "Null" Or Trim(rst2!MBMCostCenter) = Trim("") Then
                        TmpSval = "Nul"
                    Else
                        TmpSval = rst2!MBMCostCenter
                    End If
                    
                     If sValue = TmpSval Then
                        SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                        Else
                         PVCnt = PVCnt + 2
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), "O").Font.Bold = True
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), "O") = "Sub Total"
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), "P").Font.Bold = True
                         objWorksheet.Cells(8 + CDbl(PVCnt - 1), "P") = Format(SubTotal, "#,##0.00")
                         sValue = rst2!MBMCostCenter
                         SubTotal = 0
                         SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                        End If
                        
                        Diag = ""
                        PVCnt = PVCnt + 1
                        
                        Set objWorksheet = objWorkbook.Sheets(3)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                        
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "M") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "M") = TxtBordxNo
                            objWorksheet.Cells(4, "M") = rst2!BordxDate
                            objWorksheet.Cells(5, "M") = IIf(Len(SubSid), SubSid, "")
                        End If
                            Rcount = Rcount + 1
                            If Len(rst2!ReasonDesc1) Then
                             Diag = rst2!ReasonDesc1
                            End If
                            If Len(rst2!ReasonDesc2) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc2
                                 Else
                                  Diag = rst2!ReasonDesc2
                                End If
                            End If
                            If Len(rst2!ReasonDesc3) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc3
                                 Else
                                  Diag = rst2!ReasonDesc3
                                End If
                            End If
                             If Len(rst2!ReasonDesc4) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc4
                                 Else
                                  Diag = rst2!ReasonDesc4
                                End If
                            End If
                            If Len(rst2!ReasonDesc5) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc5
                                 Else
                                  Diag = rst2!ReasonDesc5
                                End If
                            End If
                            If Len(rst2!CASDiagnosis5) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!CASDiagnosis5
                                 Else
                                  Diag = rst2!CASDiagnosis5
                                End If
                            End If
                            objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Rcount
                            objWorksheet.Cells(8 + CDbl(PVCnt), "B") = rst2!CasNumber
                            objWorksheet.Cells(8 + CDbl(PVCnt), "C") = rst2!MBMNumber
                            objWorksheet.Cells(8 + CDbl(PVCnt), "D") = rst2!MBMBranch
                            objWorksheet.Cells(8 + CDbl(PVCnt), "E") = rst2!MBMDepartment
                            objWorksheet.Cells(8 + CDbl(PVCnt), "F") = rst2!MBMCostCenter
                            objWorksheet.Cells(8 + CDbl(PVCnt), "G") = rst2!MBMEmployeeNo
                            objWorksheet.Cells(8 + CDbl(PVCnt), "H") = rst2!MBMName
                            objWorksheet.Cells(8 + CDbl(PVCnt), "I") = rst2!IC
                            objWorksheet.Cells(8 + CDbl(PVCnt), "J") = rst2!PatientName
                            objWorksheet.Cells(8 + CDbl(PVCnt), "K") = CStr(Format(rst2!DateVisit, "DD/MM/YYYY"))
                            objWorksheet.Cells(8 + CDbl(PVCnt), "M") = Diag
                            objWorksheet.Cells(8 + CDbl(PVCnt), "N") = rst2!CLNType
                            If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(8 + CDbl(PVCnt), "O") = rst2!NPClinic
                            Else
                                objWorksheet.Cells(8 + CDbl(PVCnt), "O") = rst2!CLNName
                            End If
                            objWorksheet.Cells(8 + CDbl(PVCnt), "P") = rst2!CLNLocation
                            objWorksheet.Cells(8 + CDbl(PVCnt), "Q") = rst2!CLMTotalActual
                            
                            tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                            tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                       IsFirst = False
                       rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    RecFound = False
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "O").Font.Bold = True
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "O") = "Sub Total"
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "P").Font.Bold = True
                         objWorksheet.Cells(9 + CDbl(PVCnt + 1), "P") = Format(SubTotal, "#,##0.00")
                    PVCnt = PVCnt + 1
                    objWorksheet.Cells(10 + CDbl(PVCnt), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(10 + CDbl(PVCnt), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(10 + CDbl(PVCnt), "O").Font.Bold = True
                    objWorksheet.Cells(10 + CDbl(PVCnt), "O") = "Total"
                    objWorksheet.Cells(10 + CDbl(PVCnt), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(10 + CDbl(PVCnt), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(10 + CDbl(PVCnt), "P").Font.Bold = True
                    objWorksheet.Cells(10 + CDbl(PVCnt), "P") = Format(tmpSubApp, "#,##0.00")
            Call PrintCoverLetter(tmpcovercompany, tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
            End If
        ElseIf CboReportType.ListIndex = "5" Then
            Call BYSUB(objWorkbook)
            Call PrintCoverLetter(tmpcovercompany, tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
        ElseIf CboReportType.ListIndex = "2" Then
            Call BYREL(objWorkbook)
        Else
            strSql = ""
            
                    Set rst2 = GetDataFromCISWS(" BordxRefNo = '" & Trim(TxtBordxNo) & "'", 4, "CLM_PayAdviceCLNBSN")
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    RecFound = False
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    
                     strsql2 = ""
                        strsql2 = "SELECT CISDB..View_PrintBordx_Excel.Subsidiary FROM  CISDB..View_PrintBordx_Excel WHERE  BordxRefNo = '" & Trim(TxtBordxNo) & "' and CISDB..View_PrintBordx_Excel.Subsidiary is not null and ltrim(rtrim(CISDB..View_PrintBordx_Excel.Subsidiary)) <> '' GROUP BY CISDB..View_PrintBordx_Excel.Subsidiary"
                        Subsidiary.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                        Do While Not Subsidiary.EOF
                            SubSid = Subsidiary!Subsidiary
                            Subsidiary.MoveNext
                        Loop
                    
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        
                       Set objWorksheet = objWorkbook.Sheets(1)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from VIEW_OP_GRPSummaryALL where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
        
                    tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(12, "K") = BOR!BordxDate
                    objWorksheet.Cells(16, "C") = Trim(BOR!GRPCoContact)
                    objWorksheet.Cells(16, "C").WrapText = False
                    objWorksheet.Cells(10, "B").WrapText = False
                    objWorksheet.Cells(10, "B") = BOR!MBMGroupCompany
                    objWorksheet.Cells(11, "B") = BOR!GRPCoAdd1
                    objWorksheet.Cells(11, "B").WrapText = False
                    objWorksheet.Cells(12, "B") = BOR!GRPCoAdd2
                    objWorksheet.Cells(12, "B").WrapText = False
                    objWorksheet.Cells(13, "B") = BOR!GRPCoAdd3
                    objWorksheet.Cells(13, "B").WrapText = False
                    objWorksheet.Cells(14, "B") = BOR!GRPCoPhone
                    objWorksheet.Cells(14, "B").WrapText = False
                    objWorksheet.Cells(15, "B") = BOR!GRPCoFax
                    objWorksheet.Cells(15, "B").WrapText = False
                    objWorksheet.Cells(27, "B") = "OUTPATIENT MEDICAL COST FOR" & " " & BOR!MBMGroupCompany
                    objWorksheet.Cells(27, "I") = BOR!TotalVisits
                    objWorksheet.Cells(27, "K") = CDbl(BOR!CLMBordxAmt)
                    tmpcovercompany = IIf(Len(BOR!MBMGroupCompany), BOR!MBMGroupCompany, "")
                    tmpcoveraadd1 = IIf(Len(BOR!GRPCoContact), BOR!GRPCoContact, "NA")
                    tmpcoveraadd2 = IIf(Len(BOR!GRPCoAdd1), BOR!GRPCoAdd1, "NA")
                    tmpcoveraadd3 = IIf(Len(BOR!GRPCoAdd2), BOR!GRPCoAdd2, "NA")
                    tmpcoveraadd4 = IIf(Len(BOR!GRPCoAdd3), BOR!GRPCoAdd3, "NA")
                    tmpcoveraadd5 = IIf(Len(BOR!GRPCoPhone), BOR!GRPCoPhone, "NA")
                    tmpcoveraadd6 = IIf(Len(BOR!GRPCoFax), BOR!GRPCoFax, "NA")
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                    
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
                     
                        Set objWorksheet = objWorkbook.Sheets(2)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "H") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "H") = TxtBordxNo
                            objWorksheet.Cells(4, "H") = rst2!BordxDate
                            objWorksheet.Cells(5, "H") = IIf(Len(SubSid), SubSid, "")
                        End If
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            objWorksheet.Cells(16 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                            objWorksheet.Cells(16 + CDbl(PVCnt), "B") = rst2!MBMDepartment
                            objWorksheet.Cells(16 + CDbl(PVCnt), "E") = rst2!TotalCases
                            tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                            objWorksheet.Cells(16 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!TotalActual), "#,##0.00")
                            totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                            objWorksheet.Cells(16 + CDbl(PVCnt), "G") = Format(rst2!PayabletoHOS, "#,##0.00")
                            objWorksheet.Cells(16 + CDbl(PVCnt), "H") = CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS)
                            
                            tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                            tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
                        rst2.MoveNext
                    Loop
                        
                        objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(18 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(18 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                        objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(18 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(18 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                        objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
                        objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(totalMBMPaid, "#,##0.00")
                        objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
                        objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(tmpTotalHos, "#,##0.00")
                        
                    rst2.Close
                    Set rst2 = Nothing
                    
                    If RecFound = True Then
                            strSql = ""
                            Set rst2 = GetDataFromCISWS(" CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'", 1, "CLM_PayAdviceCLNBSN")
                            PVCnt = 0
                            tmpTotalActual = 0
                            tmpTotalHos = 0
                            TMPWS = 1
                            SubTotal = 0
                            IsFirst = True
                    
                        Do While Not rst2.EOF
                            Diag = ""
                            
                            If rst2!Relation = "Principle" Then
                             PVCnt = PVCnt + 1
                             SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                             Set objWorksheet = objWorkbook.Sheets(3)
                             objExcel.DisplayAlerts = False
                             Set objWorksheet = objWorkbook.Worksheets.Item(3)
                       
                                If PVCnt = 1 Then
                                    objWorksheet.Cells(2, "L") = rst2!MBMGroupCompany
                                    objWorksheet.Cells(3, "L") = TxtBordxNo
                                    objWorksheet.Cells(4, "L") = rst2!BordxDate
                                    objWorksheet.Cells(6, "L") = IIf(Len(SubSid), SubSid, "")
                                End If
                        
                                Rcount = Rcount + 1
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "A") = Rcount
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CasNumber
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = IIf(Len(rst2!MBMEmployeeNo), rst2!MBMEmployeeNo, "")
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "G") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "H") = rst2!MBMName
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "J") = rst2!MBMNumber
                                    If rst2!coverid = "00" Then
                                        objWorksheet.Cells(14 + CDbl(PVCnt), "L") = rst2!MBMName
                                    Else
                                        objWorksheet.Cells(14 + CDbl(PVCnt), "L") = rst2!MBMCName
                                    End If
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "M") = rst2!IC
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "N") = rst2!MBMDepartment
                                    If Len(rst2!ReasonDesc1) Then
                                         Diag = rst2!ReasonDesc1
                                    End If
                                    If Len(rst2!ReasonDesc2) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!ReasonDesc2
                                        Else
                                        Diag = rst2!ReasonDesc2
                                        End If
                                    End If
                                    If Len(rst2!ReasonDesc3) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!ReasonDesc3
                                        Else
                                        Diag = rst2!ReasonDesc3
                                        End If
                                    End If
                            
                                    If Len(rst2!CASDiagnosis5) Then
                                        If Len(Diag) Then
                                       Diag = Diag & " \ " & rst2!CASDiagnosis5
                                       Else
                                       Diag = rst2!CASDiagnosis5
                                        End If
                                    End If
                                                   
                                objWorksheet.Cells(14 + CDbl(PVCnt), "P") = Diag
                                If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(14 + CDbl(PVCnt), "S") = rst2!NPClinic
                                    Else
                                objWorksheet.Cells(14 + CDbl(PVCnt), "S") = rst2!CLNName
                                End If
                                objWorksheet.Cells(14 + CDbl(PVCnt), "T") = IIf(Len(rst2!MEMREL), rst2!MEMREL, "")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "U") = IIf(Len(rst2!CLNLocation), rst2!CLNLocation, "")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(rst2!DateVisit, "YYYY/MM/DD")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "W") = Format(rst2!CLMTotalActual, "#,##0.00")
                                                
                                tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                                tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMTotalActual, "#,##0.00"))
                                tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                            
                            IsFirst = False
                        End If
                        rst2.MoveNext
                    Loop
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Font.Bold = True
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V") = "Principle Sub Total"
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "W").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "W").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "W").Font.Bold = True
                         objWorksheet.Cells(14 + CDbl(PVCnt + 1), "W") = Format(SubTotal, "#,##0.00")
                        SubTotal = 0
                        Set rst2 = Nothing
                        Set rst2 = GetDataFromCISWS(" CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'", 1, "CLM_PayAdviceCLNBSN")
                        'Set rst2 = cmd.Execute
                        PVCnt = PVCnt + 2
                        Do While Not rst2.EOF
                            Diag = ""
                            If rst2!Relation <> "Principle" Then
                             PVCnt = PVCnt + 1
                             SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                             Set objWorksheet = objWorkbook.Sheets(3)
                             objExcel.DisplayAlerts = False
                             Set objWorksheet = objWorkbook.Worksheets.Item(3)
                       
                                If PVCnt = 1 Then
                                    objWorksheet.Cells(2, "L") = rst2!MBMGroupCompany
                                    objWorksheet.Cells(3, "L") = TxtBordxNo
                                    objWorksheet.Cells(4, "L") = rst2!BordxDate
                                    objWorksheet.Cells(6, "L") = IIf(Len(SubSid), SubSid, "")
                                End If
                        
                                Rcount = Rcount + 1
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "A") = Rcount
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CasNumber
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = IIf(Len(rst2!MBMEmployeeNo), rst2!MBMEmployeeNo, "")
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "G") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "H") = rst2!MBMName
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "J") = rst2!MBMNumber
                                    If rst2!coverid = "00" Then
                                        objWorksheet.Cells(14 + CDbl(PVCnt), "L") = rst2!MBMName
                                    Else
                                        objWorksheet.Cells(14 + CDbl(PVCnt), "L") = rst2!MBMCName
                                    End If
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "M") = rst2!IC
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "N") = rst2!MBMDepartment
                                    If Len(rst2!ReasonDesc1) Then
                                         Diag = rst2!ReasonDesc1
                                    End If
                                    If Len(rst2!ReasonDesc2) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!ReasonDesc2
                                        Else
                                        Diag = rst2!ReasonDesc2
                                        End If
                                    End If
                                    If Len(rst2!ReasonDesc3) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!ReasonDesc3
                                        Else
                                        Diag = rst2!ReasonDesc3
                                        End If
                                    End If
                            
                                    If Len(rst2!CASDiagnosis5) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!CASDiagnosis5
                                        Else
                                        Diag = rst2!CASDiagnosis5
                                       End If
                                    End If
                                                   
                                objWorksheet.Cells(14 + CDbl(PVCnt), "P") = Diag
                                If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(14 + CDbl(PVCnt), "S") = rst2!NPClinic
                                Else
                                objWorksheet.Cells(14 + CDbl(PVCnt), "S") = rst2!CLNName
                                End If
                                objWorksheet.Cells(14 + CDbl(PVCnt), "T") = IIf(Len(rst2!MEMREL), rst2!MEMREL, "")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "U") = IIf(Len(rst2!CLNLocation), rst2!CLNLocation, "")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(rst2!DateVisit, "YYYY/MM/DD")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "W") = Format(rst2!CLMTotalActual, "#,##0.00")
                                tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                                tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMTotalActual, "#,##0.00"))
                                tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                        End If
                        rst2.MoveNext
                    Loop
                    RecFound = False
                    objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V") = "Supplementary Sub Total"
                    
                    objWorksheet.Cells(14 + CDbl(PVCnt + 1), "W").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt + 1), "W").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt + 1), "W").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt + 1), "W") = Format(SubTotal, "#,##0.00")
                    PVCnt = PVCnt + 3
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V") = "Total"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "W").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "W").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "W").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "W") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "W") = Format(tmpsubAct, "#,##0.00")
        Call PrintCoverLetter(tmpcovercompany, tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
        End If
        End If
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=1, OpenAfterPublish:=True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        Exit Sub
End Sub
Public Sub BYDEPT(objWorkbook As Excel.Workbook)
Const METHOD_NAME = "BYDEPT"
On Error GoTo debugs

Dim strSql As String
Dim cmd As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim prm8 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim strsql2 As String
Dim PVCnt As Integer
Dim Rcount As Integer
Dim tmpTotalCases As String
Dim tmpTotalActual As String
Dim tmpTotalHos As Double
Dim tmpHOSPayee As String
Dim tmpSubMBMPaid As Double
Dim RecFound, headerstatus As Boolean
Dim tmpSeqNo As Integer
Dim totalMBMPaid As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim SubSid As String
Dim Subsidiary As New ADODB.Recordset
Dim Diag As String
Dim sqlbor As String
Dim tmpBordxMonth As String
Dim tmpBordxYear As String
Dim tmpsubAct, tmptotalcopay, tmpfinalapprove As Double
Dim tmpSubApp As Double
Dim SubTotal As Double
Dim SubTotal_App As Double
Dim SubTotal_Copay As Double
Dim IsFirst As Boolean
Dim sValue As String
headerstatus = False
strSql = ""
                    strSql = " BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd4.ActiveConnection = MediConnCISWS
                    cmd4.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd4.CommandType = adCmdStoredProc
                    cmd4.CommandTimeout = 300
                    Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strSql)
                    cmd4.Parameters.Append Prm7
                    Set prm8 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 6)
                    cmd4.Parameters.Append prm8
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd4.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    RecFound = False
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    
                     strsql2 = ""
                        strsql2 = "SELECT CISDB..View_PrintBordx_Excel.Subsidiary FROM  CISDB..View_PrintBordx_Excel WHERE  BordxRefNo = '" & Trim(TxtBordxNo) & "' and CISDB..View_PrintBordx_Excel.Subsidiary is not null and ltrim(rtrim(CISDB..View_PrintBordx_Excel.Subsidiary)) <> '' GROUP BY CISDB..View_PrintBordx_Excel.Subsidiary"
                        Subsidiary.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                        Do While Not Subsidiary.EOF
                            SubSid = Subsidiary!Subsidiary
                            Subsidiary.MoveNext
                        Loop
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        'objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "F") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "F") = TxtBordxNo
                            objWorksheet.Cells(4, "F") = rst2!BordxDate
                            objWorksheet.Cells(5, "F") = IIf(Len(SubSid), SubSid, "")
                        End If
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "B") = rst2!MBMDeptCode
                            objWorksheet.Cells(8 + CDbl(PVCnt), "C") = rst2!MBMDepartment
                            objWorksheet.Cells(8 + CDbl(PVCnt), "E") = rst2!TotalCases
                            tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                            objWorksheet.Cells(8 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!TotalActual), "#,##0.00")
                            totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                            objWorksheet.Cells(8 + CDbl(PVCnt), "G") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                            tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                            tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
                        rst2.MoveNext
                    Loop
                        objWorksheet.Cells(10 + CDbl(PVCnt), "D").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "D").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "D") = "Total"
                        
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                        
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                    rst2.Close
                    Set rst2 = Nothing
                
                Set objWorksheet = objWorkbook.Sheets(1)
                    'objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from VIEW_OP_GRPSummaryALL where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
        
                    tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(12, "K") = BOR!BordxDate
                    objWorksheet.Cells(16, "C") = Trim(BOR!GRPCoContact)
                    objWorksheet.Cells(16, "C").WrapText = False
                    objWorksheet.Cells(10, "B").WrapText = False
                    objWorksheet.Cells(10, "B") = BOR!MBMGroupCompany
                    objWorksheet.Cells(11, "B") = BOR!GRPCoAdd1
                    objWorksheet.Cells(11, "B").WrapText = False
                    objWorksheet.Cells(12, "B") = BOR!GRPCoAdd2
                    objWorksheet.Cells(12, "B").WrapText = False
                    objWorksheet.Cells(13, "B") = BOR!GRPCoAdd3
                    objWorksheet.Cells(13, "B").WrapText = False
                    objWorksheet.Cells(14, "B") = BOR!GRPCoPhone
                    objWorksheet.Cells(14, "B").WrapText = False
                    objWorksheet.Cells(15, "B") = BOR!GRPCoFax
                    objWorksheet.Cells(15, "B").WrapText = False
                    objWorksheet.Cells(27, "B") = "OUTPATIENT MEDICAL COST FOR" & " " & BOR!MBMGroupCompany
                    
                    
                    tmpcovercompany = IIf(Len(BOR!MBMGroupCompany), BOR!MBMGroupCompany, "")
                    tmpcoveraadd1 = IIf(Len(BOR!GRPCoContact), BOR!GRPCoContact, "NA")
                    tmpcoveraadd2 = IIf(Len(BOR!GRPCoAdd1), BOR!GRPCoAdd1, "NA")
                    tmpcoveraadd3 = IIf(Len(BOR!GRPCoAdd2), BOR!GRPCoAdd2, "NA")
                    tmpcoveraadd4 = IIf(Len(BOR!GRPCoAdd3), BOR!GRPCoAdd3, "NA")
                    tmpcoveraadd5 = IIf(Len(BOR!GRPCoPhone), BOR!GRPCoPhone, "NA")
                    tmpcoveraadd6 = IIf(Len(BOR!GRPCoFax), BOR!GRPCoFax, "NA")
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                    
               
                objWorksheet.Cells(27, "I") = BOR!TotalVisits
                objWorksheet.Cells(27, "K") = CDbl(BOR!CLMBordxAmt)
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
                
                If RecFound = True Then
                    strSql = ""
                    strSql = " CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 300
                    Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, strSql)
                    cmd2.Parameters.Append Prm3
                    Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 3000, 1)
                    cmd2.Parameters.Append prm4
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd2.Execute
                    PVCnt = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    TMPWS = 1
                    SubTotal = 0
                    SubTotal_App = 0
                    SubTotal_Copay = 0
                    IsFirst = True
                    
                    Do While Not rst2.EOF
                    
                    Diag = ""
                    If IsFirst Then
                    sValue = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                    End If
                    
                     If sValue = rst2!MBMDepartment Then
                        SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                        SubTotal_Copay = SubTotal_Copay + CDbl(rst2!CLMCoPayColl)
                        SubTotal_App = SubTotal_App + CDbl(rst2!CLMFinalApprove)
                     Else
                         PVCnt = PVCnt + 2
                         
                         tmpSeqNo = tmpSeqNo - 2
                         
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "N").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "N") = "Sub Total"
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "O").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "O") = Format(SubTotal, "#,##0.00")
                         
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "P").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "P") = Format(SubTotal_Copay, "#,##0.00")
                         
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "Q").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "Q").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "Q").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "Q") = Format(SubTotal_App, "#,##0.00")
                         
                         
                         sValue = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                         SubTotal = 0
                         SubTotal_Copay = 0
                         SubTotal_App = 0
                         SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                         SubTotal_Copay = SubTotal_Copay + CDbl(rst2!CLMCoPayColl)
                         SubTotal_App = SubTotal_App + CDbl(rst2!CLMFinalApprove)
                    End If
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(3)
                        'objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                        If headerstatus = False Then
                            objWorksheet.Cells(2, "I") = ": " & rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "I") = ": " & TxtBordxNo
                            objWorksheet.Cells(4, "I") = ": " & rst2!BordxDate
                            objWorksheet.Cells(5, "I") = ": " & IIf(Len(SubSid), SubSid, "")
                            headerstatus = True
                        End If
                        
                            Rcount = Rcount + 1
                            objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Rcount
                            objWorksheet.Cells(7 + CDbl(PVCnt), "B") = rst2!CasNumber
                            objWorksheet.Cells(7 + CDbl(PVCnt), "C") = IIf(Len(rst2!MBMEmployeeNo), rst2!MBMEmployeeNo, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "D") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "E") = rst2!MBMName
                            objWorksheet.Cells(7 + CDbl(PVCnt), "F") = rst2!MBMNumber
                            If rst2!coverid = "00" Then
                                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = rst2!MBMName
                            Else
                                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = rst2!MBMCName
                            End If
                            
                            objWorksheet.Cells(7 + CDbl(PVCnt), "H") = rst2!IC
                            objWorksheet.Cells(7 + CDbl(PVCnt), "I") = IIf(Len(rst2!MBMDeptCode), rst2!MBMDeptCode, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "J") = rst2!MBMDepartment
                            'objWorksheet.Cells(14 + CDbl(PVCnt), "O") = rst2!MBMCostcenter
                            
                            If Len(rst2!ReasonDesc1) Then
                             Diag = rst2!ReasonDesc1
                            End If
                            
                            If Len(rst2!ReasonDesc2) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc2
                                 Else
                                  Diag = rst2!ReasonDesc2
                                End If
                            End If
                            If Len(rst2!ReasonDesc3) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc3
                                 Else
                                  Diag = rst2!ReasonDesc3
                                End If
                            End If
                            If Len(rst2!CASDiagnosis5) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!CASDiagnosis5
                                Else
                                 Diag = rst2!CASDiagnosis5
                               End If
                            End If
                            objWorksheet.Cells(7 + CDbl(PVCnt), "K") = Diag
                            If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(7 + CDbl(PVCnt), "L") = rst2!NPClinic
                            Else
                                objWorksheet.Cells(7 + CDbl(PVCnt), "L") = rst2!CLNName
                            End If
                            
                           objWorksheet.Cells(7 + CDbl(PVCnt), "M") = IIf(Len(rst2!CLNLocation), rst2!CLNLocation, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "N") = Format(rst2!DateVisit, "YYYY/MM/DD")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "O") = Format(rst2!CLMTotalActual, "#,##0.00")
                            tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                            tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                            objWorksheet.Cells(7 + CDbl(PVCnt), "P") = Format(rst2!CLMCoPayColl, "#,##0.00")
                            tmptotalcopay = tmptotalcopay + CDbl(rst2!CLMCoPayColl)
                            objWorksheet.Cells(7 + CDbl(PVCnt), "Q") = Format(rst2!CLMFinalApprove, "#,##0.00")
                            tmpfinalapprove = tmpfinalapprove + CDbl(rst2!CLMFinalApprove)
                            tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMTotalActual, "#,##0.00"))
                            
                        
                        IsFirst = False
                        rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    RecFound = False
                            
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "N").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "N") = "Sub Total"
                      
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "O").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "O") = Format(SubTotal, "#,##0.00")
                      
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "P").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "P") = Format(SubTotal_Copay, "#,##0.00")
                     
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "Q").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "Q").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "Q").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "Q") = Format(SubTotal_App, "#,##0.00")
                    
            
                    PVCnt = PVCnt + 4
                    
                    objWorksheet.Cells(7 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(7 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(7 + CDbl(PVCnt), "N").Font.Bold = True
                    objWorksheet.Cells(7 + CDbl(PVCnt), "N") = "Total"
                    
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O").Font.Bold = True
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O") = Format(tmpsubAct, "#,##0.00")
                    
                    objWorksheet.Cells(7 + CDbl(PVCnt), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(7 + CDbl(PVCnt), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(7 + CDbl(PVCnt), "P").Font.Bold = True
                    objWorksheet.Cells(7 + CDbl(PVCnt), "P") = Format(tmptotalcopay, "#,##0.00")
                    
                    objWorksheet.Cells(7 + CDbl(PVCnt), "Q").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(7 + CDbl(PVCnt), "Q").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(7 + CDbl(PVCnt), "Q").Font.Bold = True
                    objWorksheet.Cells(7 + CDbl(PVCnt), "Q") = Format(tmpfinalapprove, "#,##0.00")
                    
                    End If
                    
                    
 
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
Public Sub BYREL(objWorkbook As Excel.Workbook)
Dim strSql As String
Dim cmd As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim prm8 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim strsql2 As String
Dim PVCnt As Integer
Dim Rcount As Integer
Dim tmpTotalCases As String
Dim tmpTotalActual As String
Dim tmpTotalHos As Double
Dim tmpHOSPayee As String
Dim tmpSubMBMPaid As Double
Dim RecFound As Boolean
Dim tmpSeqNo As Integer
Dim totalMBMPaid As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim SubSid As String
Dim Subsidiary As New ADODB.Recordset
'Dim objWorksheet As Excel.Worksheet
Dim Diag As String
Dim sqlbor As String
Dim tmpBordxMonth As String
Dim tmpBordxYear As String
Dim tmpsubAct, tmpsubInvoiceAmt, tmpsubcopay As Double
Dim tmpSubApp As Double
Dim SubTotal, SubInvoiceTotal, SubcOPAYTotal As Double
Dim IsFirst As Boolean
Dim sValue As String
strSql = ""
tmpsubAct = 0
tmpsubInvoiceAmt = 0
tmpsubcopay = 0
            
                    Set rst2 = GetDataFromCISWS(" BordxRefNo = '" & Trim(TxtBordxNo) & "'", 12, "CLM_PayAdviceCLNBSN")
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    RecFound = False
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    
                     strsql2 = ""
                        strsql2 = "SELECT CISDB..View_PrintBordx_Excel.Subsidiary FROM  CISDB..View_PrintBordx_Excel WHERE  BordxRefNo = '" & Trim(TxtBordxNo) & "' and CISDB..View_PrintBordx_Excel.Subsidiary is not null and ltrim(rtrim(CISDB..View_PrintBordx_Excel.Subsidiary)) <> '' GROUP BY CISDB..View_PrintBordx_Excel.Subsidiary"
                        Subsidiary.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                        Do While Not Subsidiary.EOF
                            SubSid = Subsidiary!Subsidiary
                            Subsidiary.MoveNext
                        Loop
                    
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        
                       Set objWorksheet = objWorkbook.Sheets(1)
                    'objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from VIEW_OP_GRPSummaryALL where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
        
                    tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(12, "K") = BOR!BordxDate
                    objWorksheet.Cells(16, "C") = Trim(BOR!GRPCoContact)
                    objWorksheet.Cells(16, "C").WrapText = False
                    objWorksheet.Cells(10, "B").WrapText = False
                    objWorksheet.Cells(10, "B") = BOR!MBMGroupCompany
                    objWorksheet.Cells(11, "B") = BOR!GRPCoAdd1
                    objWorksheet.Cells(11, "B").WrapText = False
                    objWorksheet.Cells(12, "B") = BOR!GRPCoAdd2
                    objWorksheet.Cells(12, "B").WrapText = False
                    objWorksheet.Cells(13, "B") = BOR!GRPCoAdd3
                    objWorksheet.Cells(13, "B").WrapText = False
                    objWorksheet.Cells(14, "B") = BOR!GRPCoPhone
                    objWorksheet.Cells(14, "B").WrapText = False
                    objWorksheet.Cells(15, "B") = BOR!GRPCoFax
                    objWorksheet.Cells(15, "B").WrapText = False
                    objWorksheet.Cells(27, "B") = "OUTPATIENT MEDICAL COST FOR" & " " & BOR!MBMGroupCompany
                    objWorksheet.Cells(27, "I") = BOR!TotalVisits
                    objWorksheet.Cells(27, "K") = CDbl(BOR!CLMBordxAmt)
                    tmpcovercompany = IIf(Len(BOR!MBMGroupCompany), BOR!MBMGroupCompany, "")
                    tmpcoveraadd1 = IIf(Len(BOR!GRPCoContact), BOR!GRPCoContact, "NA")
                    tmpcoveraadd2 = IIf(Len(BOR!GRPCoAdd1), BOR!GRPCoAdd1, "NA")
                    tmpcoveraadd3 = IIf(Len(BOR!GRPCoAdd2), BOR!GRPCoAdd2, "NA")
                    tmpcoveraadd4 = IIf(Len(BOR!GRPCoAdd3), BOR!GRPCoAdd3, "NA")
                    tmpcoveraadd5 = IIf(Len(BOR!GRPCoPhone), BOR!GRPCoPhone, "NA")
                    tmpcoveraadd6 = IIf(Len(BOR!GRPCoFax), BOR!GRPCoFax, "NA")
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                    
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
                     
                        Set objWorksheet = objWorkbook.Sheets(2)
                        'objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "E") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "E") = TxtBordxNo
                            objWorksheet.Cells(4, "E") = rst2!BordxDate
                            objWorksheet.Cells(5, "E") = IIf(Len(SubSid), SubSid, "")
                        End If
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "B") = rst2!MBMDeptCode
                            objWorksheet.Cells(8 + CDbl(PVCnt), "C") = rst2!MBMDepartment
                            objWorksheet.Cells(8 + CDbl(PVCnt), "D") = rst2!TotalCases
                            tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                            objWorksheet.Cells(8 + CDbl(PVCnt), "E") = Format(rst2!ClaimAmt, "#,##0.00")
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!ClaimAmt), "#,##0.00")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "F") = rst2!CLNType
                            rst2.MoveNext
                    Loop
                        objWorksheet.Cells(10 + CDbl(PVCnt), "D").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "D").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "D") = Format(tmpTotalCases, "#,##0.00")
                        
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E") = Format(tmpTotalActual, "#,##0.00")
                    rst2.Close
                    Set rst2 = Nothing
                    
                    If RecFound = True Then
                            strSql = ""
                            Set rst2 = GetDataFromCISWS(" CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'", 1, "CLM_PayAdviceCLNBSN")
                            PVCnt = 0
                            tmpTotalActual = 0
                            tmpTotalHos = 0
                            TMPWS = 1
                            SubTotal = 0
                            SubcOPAYTotal = 0
                            SubInvoiceTotal = 0
                            IsFirst = True
                    
                        Do While Not rst2.EOF
                            Diag = ""
                            If rst2!Relation = "Principle" Then
                             PVCnt = PVCnt + 1
                             SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                             Set objWorksheet = objWorkbook.Sheets(3)
                             Set objWorksheet = objWorkbook.Worksheets.Item(3)
                                If PVCnt = 1 Then
                                    objWorksheet.Cells(2, "G") = rst2!MBMGroupCompany
                                    objWorksheet.Cells(3, "G") = TxtBordxNo
                                    objWorksheet.Cells(4, "G") = rst2!BordxDate
                                    objWorksheet.Cells(5, "G") = IIf(Len(SubSid), SubSid, "")
                                End If
                        
                                Rcount = Rcount + 1
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Rcount
                                    '--B--Claim No
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "B") = rst2!CasNumber
                                    '--C--Visit Type
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "C") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                                    '--D--Staff Name
                                     objWorksheet.Cells(8 + CDbl(PVCnt), "D") = rst2!MBMName
                                     '--E--MBM Number
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "E") = rst2!MBMNumber
                                    '--F--Patient Name
                                    If rst2!coverid = "00" Then
                                        objWorksheet.Cells(8 + CDbl(PVCnt), "F") = rst2!MBMName
                                    Else
                                        objWorksheet.Cells(8 + CDbl(PVCnt), "F") = rst2!MBMCName
                                    End If
                                    '--H--IC Number
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "H") = rst2!IC
                                    '--G--Relation Code
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "G") = IIf(Len(rst2!MEMREL), rst2!MEMREL, "")
                                    '--I--Dept
                                     objWorksheet.Cells(8 + CDbl(PVCnt), "I") = rst2!MBMDepartment
                                    '--J--Diagnosis
                                     If Len(rst2!ReasonDesc1) Then
                                         Diag = rst2!ReasonDesc1
                                    End If
                                    If Len(rst2!ReasonDesc2) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!ReasonDesc2
                                        Else
                                        Diag = rst2!ReasonDesc2
                                        End If
                                    End If
                                    If Len(rst2!ReasonDesc3) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!ReasonDesc3
                                        Else
                                        Diag = rst2!ReasonDesc3
                                        End If
                                    End If
                            
                                    If Len(rst2!CASDiagnosis5) Then
                                        If Len(Diag) Then
                                       Diag = Diag & " \ " & rst2!CASDiagnosis5
                                       Else
                                       Diag = rst2!CASDiagnosis5
                                        End If
                                    End If
                                   objWorksheet.Cells(8 + CDbl(PVCnt), "J") = Diag
                                    '--K--Clinic
                                If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(8 + CDbl(PVCnt), "K") = rst2!NPClinic
                                    Else
                                objWorksheet.Cells(8 + CDbl(PVCnt), "K") = rst2!CLNName
                                End If
                                '--L--Location
                                objWorksheet.Cells(8 + CDbl(PVCnt), "L") = IIf(Len(rst2!CLNLocation), rst2!CLNLocation, "")
                                    '--M--Visit Date
                                objWorksheet.Cells(8 + CDbl(PVCnt), "M") = Format(rst2!DateVisit, "YYYY/MM/DD")
                                 '--N--Claim Amt
                                objWorksheet.Cells(8 + CDbl(PVCnt), "N") = Format(rst2!CLMTotalActual, "#,##0.00")
                                '--O--Co Pay
                                objWorksheet.Cells(8 + CDbl(PVCnt), "O") = Format(rst2!CLMCoPayCLNColl, "#,##0.00")
                                SubcOPAYTotal = SubcOPAYTotal + CDbl(rst2!CLMCoPayCLNColl)
                                '--P--Invoice Amt
                                objWorksheet.Cells(8 + CDbl(PVCnt), "P") = Format(rst2!CLMBordxAmt, "#,##0.00")
                                SubInvoiceTotal = SubInvoiceTotal + CDbl(rst2!CLMBordxAmt)
                                
                                tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                                tmpsubcopay = CDbl(tmpsubcopay) + CDbl(rst2!CLMCoPayCLNColl)
                                tmpsubInvoiceAmt = CDbl(tmpsubInvoiceAmt) + CDbl(rst2!CLMBordxAmt)
                                
                                tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMTotalActual, "#,##0.00"))
                                tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                            
                            IsFirst = False
                        End If
                        rst2.MoveNext
                    Loop
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "M").Font.Bold = True
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "M") = "Principle Sub Total"
                         
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "N").Font.Bold = True
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "N") = Format(SubTotal, "#,##0.00")
                         
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "O").Font.Bold = True
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "O") = Format(SubcOPAYTotal, "#,##0.00")
                         
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "P").Font.Bold = True
                         objWorksheet.Cells(8 + CDbl(PVCnt + 1), "P") = Format(SubInvoiceTotal, "#,##0.00")
                        SubTotal = 0
                        SubcOPAYTotal = 0
                        SubInvoiceTotal = 0
                        Set rst2 = Nothing
                        Set rst2 = GetDataFromCISWS(" CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'", 1, "CLM_PayAdviceCLNBSN")
                        'Set rst2 = cmd.Execute
                        PVCnt = PVCnt + 2
                        Do While Not rst2.EOF
                            Diag = ""
                            If rst2!Relation <> "Principle" Then
                             PVCnt = PVCnt + 1
                             SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                             Set objWorksheet = objWorkbook.Sheets(3)
                             Set objWorksheet = objWorkbook.Worksheets.Item(3)
                       
                                If PVCnt = 1 Then
                                    objWorksheet.Cells(2, "G") = rst2!MBMGroupCompany
                                    objWorksheet.Cells(3, "G") = TxtBordxNo
                                    objWorksheet.Cells(4, "G") = rst2!BordxDate
                                    objWorksheet.Cells(5, "G") = IIf(Len(SubSid), SubSid, "")
                                End If
                        
                                Rcount = Rcount + 1
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Rcount
                                    '--B--Claim No
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "B") = rst2!CasNumber
                                    '--C--Visit Type
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "C") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                                    '--D--Staff Name
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "D") = rst2!MBMName
                                    '--E--MBM Number
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "E") = rst2!MBMNumber
                                    '--F--Patient Name
                                    If rst2!coverid = "00" Then
                                        objWorksheet.Cells(8 + CDbl(PVCnt), "F") = rst2!MBMName
                                    Else
                                        objWorksheet.Cells(8 + CDbl(PVCnt), "F") = rst2!MBMCName
                                    End If
                                    '--G--Relation Code
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "G") = IIf(Len(rst2!MEMREL), rst2!MEMREL, "")
                                    '--H--IC Number
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "H") = rst2!IC
                                    '--I--Dept
                                    objWorksheet.Cells(8 + CDbl(PVCnt), "I") = rst2!MBMDepartment
                                    If Len(rst2!ReasonDesc1) Then
                                         Diag = rst2!ReasonDesc1
                                    End If
                                    If Len(rst2!ReasonDesc2) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!ReasonDesc2
                                        Else
                                        Diag = rst2!ReasonDesc2
                                        End If
                                    End If
                                    If Len(rst2!ReasonDesc3) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!ReasonDesc3
                                        Else
                                        Diag = rst2!ReasonDesc3
                                        End If
                                    End If
                            
                                    If Len(rst2!CASDiagnosis5) Then
                                        If Len(Diag) Then
                                        Diag = Diag & " \ " & rst2!CASDiagnosis5
                                        Else
                                        Diag = rst2!CASDiagnosis5
                                       End If
                                    End If
                                 '--J--Diagnosis
                                objWorksheet.Cells(8 + CDbl(PVCnt), "J") = Diag
                                '--K--Clinic
                                If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(8 + CDbl(PVCnt), "K") = rst2!NPClinic
                                Else
                                objWorksheet.Cells(8 + CDbl(PVCnt), "K") = rst2!CLNName
                                End If
                                '--L--Location
                                objWorksheet.Cells(8 + CDbl(PVCnt), "L") = IIf(Len(rst2!CLNLocation), rst2!CLNLocation, "")
                                '--M--Visit Date
                                objWorksheet.Cells(8 + CDbl(PVCnt), "M") = Format(rst2!DateVisit, "YYYY/MM/DD")
                                '--N--Claim Amt
                                objWorksheet.Cells(8 + CDbl(PVCnt), "N") = Format(rst2!CLMTotalActual, "#,##0.00")
                                '--O--Co Pay
                                objWorksheet.Cells(8 + CDbl(PVCnt), "O") = Format(rst2!CLMCoPayCLNColl, "#,##0.00")
                                SubcOPAYTotal = SubcOPAYTotal + CDbl(rst2!CLMCoPayCLNColl)
                                '--P--Invoice Amt
                                objWorksheet.Cells(8 + CDbl(PVCnt), "P") = Format(rst2!CLMBordxAmt, "#,##0.00")
                                SubInvoiceTotal = SubInvoiceTotal + CDbl(rst2!CLMBordxAmt)
                                
                                tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                                tmpsubcopay = CDbl(tmpsubcopay) + CDbl(rst2!CLMCoPayCLNColl)
                                tmpsubInvoiceAmt = CDbl(tmpsubInvoiceAmt) + CDbl(rst2!CLMBordxAmt)
                                
                                tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMTotalActual, "#,##0.00"))
                                tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                        End If
                        rst2.MoveNext
                    Loop
                    RecFound = False
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "M").Font.Bold = True
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "M") = "Supplementary Sub Total"
                    
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "N").Font.Bold = True
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "N") = Format(SubTotal, "#,##0.00")
                    
                    
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "O").Font.Bold = True
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "O") = Format(SubcOPAYTotal, "#,##0.00")
                    
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "P").Font.Bold = True
                    objWorksheet.Cells(8 + CDbl(PVCnt + 1), "P") = Format(SubInvoiceTotal, "#,##0.00")
                    
                    
                    PVCnt = PVCnt + 3
                    objWorksheet.Cells(8 + CDbl(PVCnt), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(8 + CDbl(PVCnt), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(8 + CDbl(PVCnt), "M").Font.Bold = True
                    objWorksheet.Cells(8 + CDbl(PVCnt), "M") = "Total"
                    
                    objWorksheet.Cells(8 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(8 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(8 + CDbl(PVCnt), "N").Font.Bold = True
                    objWorksheet.Cells(8 + CDbl(PVCnt), "N") = Format(tmpsubAct, "#,##0.00")
                    
                    objWorksheet.Cells(8 + CDbl(PVCnt), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(8 + CDbl(PVCnt), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(8 + CDbl(PVCnt), "O").Font.Bold = True
                    objWorksheet.Cells(8 + CDbl(PVCnt), "O") = Format(tmpsubcopay, "#,##0.00")
                    
                    objWorksheet.Cells(8 + CDbl(PVCnt), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(8 + CDbl(PVCnt), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(8 + CDbl(PVCnt), "P").Font.Bold = True
                    objWorksheet.Cells(8 + CDbl(PVCnt), "P") = Format(tmpsubInvoiceAmt, "#,##0.00")
                    'objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpsubAct, "#,##0.00")
        Call PrintCoverLetter(tmpcovercompany, tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
        End If
End Sub


Public Sub BYSUB(objWorkbook As Excel.Workbook)
Const METHOD_NAME = "BYSUB"
On Error GoTo debugs


Dim strSql As String
Dim cmd As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim prm8 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim strsql2 As String
Dim PVCnt As Integer
Dim Rcount As Integer
Dim tmpTotalCases As String
Dim tmpTotalActual As String
Dim tmpTotalHos As Double
Dim tmpHOSPayee As String
Dim tmpSubMBMPaid As Double
Dim RecFound As Boolean
Dim tmpSeqNo As Integer
Dim totalMBMPaid As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim SubSid As String
Dim Subsidiary As New ADODB.Recordset
'Dim objWorksheet As Excel.Worksheet
Dim Diag As String
Dim sqlbor As String
Dim tmpBordxMonth As String
Dim tmpBordxYear As String
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim SubTotal As Double
Dim IsFirst As Boolean
Dim sValue As String


 strSql = ""
                    strSql = " BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd4.ActiveConnection = MediConnCISWS
                    cmd4.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd4.CommandType = adCmdStoredProc
                    cmd4.CommandTimeout = 300
                    Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strSql)
                    cmd4.Parameters.Append Prm7
                    Set prm8 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 6)
                    cmd4.Parameters.Append prm8
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd4.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    RecFound = False
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    
                     strsql2 = ""
                        strsql2 = "SELECT CISDB..View_PrintBordx_Excel.Subsidiary FROM  CISDB..View_PrintBordx_Excel WHERE  BordxRefNo = '" & Trim(TxtBordxNo) & "' and CISDB..View_PrintBordx_Excel.Subsidiary is not null and ltrim(rtrim(CISDB..View_PrintBordx_Excel.Subsidiary)) <> '' GROUP BY CISDB..View_PrintBordx_Excel.Subsidiary"
                        Subsidiary.Open strsql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                        Do While Not Subsidiary.EOF
                            SubSid = Subsidiary!Subsidiary
                            Subsidiary.MoveNext
                        Loop
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        'objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "F") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "F") = TxtBordxNo
                            objWorksheet.Cells(4, "F") = rst2!BordxDate
                            objWorksheet.Cells(5, "F") = IIf(Len(SubSid), SubSid, "")
                        End If
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "B") = rst2!MBMDeptCode
                            objWorksheet.Cells(8 + CDbl(PVCnt), "C") = rst2!MBMDepartment
                            objWorksheet.Cells(8 + CDbl(PVCnt), "E") = rst2!TotalCases
                            tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                            objWorksheet.Cells(8 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!TotalActual), "#,##0.00")
                            totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                            objWorksheet.Cells(8 + CDbl(PVCnt), "G") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                            tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                            tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
                        rst2.MoveNext
                    Loop
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                        'objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
                        'objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
                        'objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
                        'objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(tmpTotalHos, "#,##0.00")
                    rst2.Close
                    Set rst2 = Nothing
                
                Set objWorksheet = objWorkbook.Sheets(1)
                    'objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from VIEW_OP_SUBSummaryALL where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
        
                    tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(12, "K") = BOR!BordxDate
                    objWorksheet.Cells(16, "C") = BOR!PIC
                    objWorksheet.Cells(10, "B") = BOR!SubsidaryName
                    objWorksheet.Cells(11, "B") = BOR!SubAdd1
                    objWorksheet.Cells(12, "B") = BOR!SubAdd2
                    objWorksheet.Cells(13, "B") = BOR!SubAdd3
                    'objWorksheet.Cells(18, "C") = BOR!Postcode
                    objWorksheet.Cells(14, "B") = BOR!PhoneNumber
                    'objWorksheet.Cells(20, "C") = BOR!GRPCoFax
                    objWorksheet.Cells(27, "B") = "OUTPATIENT MEDICAL COST FOR" & " " & BOR!SubsidaryName
                    objWorksheet.Cells(27, "I") = BOR!TotalVisits
                    objWorksheet.Cells(27, "K") = CDbl(BOR!CLMBordxAmt)
                
                    tmpcovercompany = IIf(Len(BOR!SubsidaryName), BOR!SubsidaryName, "")
                    tmpcoveraadd1 = IIf(Len(BOR!PIC), BOR!PIC, "NA")
                    tmpcoveraadd2 = IIf(Len(BOR!SubAdd1), BOR!SubAdd1, "NA")
                    tmpcoveraadd3 = IIf(Len(BOR!SubAdd2), BOR!SubAdd2, "NA")
                    tmpcoveraadd4 = IIf(Len(BOR!SubAdd3), BOR!SubAdd3, "NA")
                    tmpcoveraadd5 = IIf(Len(BOR!PhoneNumber), BOR!PhoneNumber, "NA")
                    tmpcoveraadd6 = "NA"
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
                
                If RecFound = True Then
                    strSql = ""
                    strSql = " CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 300
                    Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, strSql)
                    cmd2.Parameters.Append Prm3
                    Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 3000, 1)
                    cmd2.Parameters.Append prm4
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd2.Execute
                    PVCnt = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    TMPWS = 1
                    SubTotal = 0
                    IsFirst = True
                    
                    Do While Not rst2.EOF
                    
                    Diag = ""
                    If IsFirst Then
                    sValue = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                    End If
                    
                     If sValue = rst2!MBMDepartment Then
                        SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                     Else
                         PVCnt = PVCnt + 2
                         
                         tmpSeqNo = tmpSeqNo - 2
                         
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "N").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "N") = "Sub Total"
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "O").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt - 1), "O") = Format(SubTotal, "#,##0.00")
                         sValue = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                         SubTotal = 0
                         SubTotal = SubTotal + CDbl(rst2!CLMTotalActual)
                    End If
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(3)
                        'objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "G") = ": " & rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "G") = ": " & TxtBordxNo
                            objWorksheet.Cells(4, "G") = ": " & rst2!BordxDate
                            objWorksheet.Cells(5, "G") = ": " & IIf(Len(SubSid), SubSid, "")
                        End If
                        
                            Rcount = Rcount + 1
                            objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Rcount
                            objWorksheet.Cells(7 + CDbl(PVCnt), "B") = rst2!CasNumber
                            objWorksheet.Cells(7 + CDbl(PVCnt), "C") = IIf(Len(rst2!MBMEmployeeNo), rst2!MBMEmployeeNo, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "D") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "E") = rst2!MBMName
                            objWorksheet.Cells(7 + CDbl(PVCnt), "F") = rst2!MBMNumber
                            If rst2!coverid = "00" Then
                                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = rst2!MBMName
                            Else
                                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = rst2!MBMCName
                            End If
                            
                            objWorksheet.Cells(7 + CDbl(PVCnt), "H") = rst2!IC
                            objWorksheet.Cells(7 + CDbl(PVCnt), "I") = rst2!MBMDepartment
                            'objWorksheet.Cells(14 + CDbl(PVCnt), "O") = rst2!MBMCostcenter
                           ' objWorksheet.Cells(14 + CDbl(PVCnt), "P") = IIf(Len(rst2!MBMDeptCode), rst2!MBMDeptCode, "")
                            If Len(rst2!ReasonDesc1) Then
                             Diag = rst2!ReasonDesc1
                            End If
                            
                            If Len(rst2!ReasonDesc2) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc2
                                 Else
                                  Diag = rst2!ReasonDesc2
                                End If
                            End If
                            If Len(rst2!ReasonDesc3) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!ReasonDesc3
                                 Else
                                  Diag = rst2!ReasonDesc3
                                End If
                            End If
                            If Len(rst2!CASDiagnosis5) Then
                                If Len(Diag) Then
                                 Diag = Diag & " \ " & rst2!CASDiagnosis5
                                Else
                                 Diag = rst2!CASDiagnosis5
                               End If
                            End If
                            objWorksheet.Cells(7 + CDbl(PVCnt), "J") = Diag
                            If rst2!ClinicCode = "NP" Or rst2!ClinicCode = "NPS" Then
                                objWorksheet.Cells(7 + CDbl(PVCnt), "K") = rst2!NPClinic
                            Else
                                objWorksheet.Cells(7 + CDbl(PVCnt), "K") = rst2!CLNName
                            End If
                            objWorksheet.Cells(7 + CDbl(PVCnt), "L") = IIf(Len(rst2!CLNLocation), rst2!CLNLocation, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "M") = rst2!MCDays
                            objWorksheet.Cells(7 + CDbl(PVCnt), "N") = Format(rst2!DateVisit, "YYYY/MM/DD")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "O") = Format(rst2!CLMTotalActual, "#,##0.00")
                  
                            tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                            tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMTotalActual, "#,##0.00"))
                            tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                  
                        IsFirst = False
                        rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    RecFound = False
                    

                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "N").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "N") = "Sub Total"
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "O").Font.Bold = True
                         objWorksheet.Cells(7 + CDbl(PVCnt + 1), "O") = Format(SubTotal, "#,##0.00")
                         
                    PVCnt = PVCnt + 2
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O").Font.Bold = True
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(7 + CDbl(PVCnt), "O") = Format(tmpsubAct, "#,##0.00")
            End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub PrintTemplateSummary()
Const METHOD_NAME = "PrintTemplateSummary"
On Error GoTo debugs

Dim StrSql1 As String
Dim strsqlPlutoBor As String
Dim PlutoBor As New ADODB.Recordset
Dim sqlinsert As String
Dim sqlbor As String
Dim tmpbordxdate As String
Dim tmpDateVisit As String
Dim TempMBMName As String
Dim TempMBMDept As String
Dim tempPath As String
Dim PLUTOStatus As Boolean
Dim BOR As New ADODB.Recordset
Dim TempRemarks As String
Dim tempGrpCompany As String
Dim TempCLNName As String
Dim TempMBMCName As String
Dim TempPatientName As String
Dim TempCASDiagnosis5 As String
Dim PVCnt As Integer
Dim tmpSeqNo As Integer
Dim tmpTotalBordxAmt As String
Dim tmpTotalActual As String
        tmpTotalBordxAmt = 0
        PVCnt = 0
        tmpSeqNo = 0
        tmpTotalActual = 0
            Set objWorksheet = objWorkbook.Sheets(2)
        'GP
        sqlbor = "Select * from VIEW_CISMBMInquiry_GPSummary where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
        BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not BOR.EOF
            'RecFound = True
            PVCnt = PVCnt + 1
                If PVCnt = "1" Then
                    
                    objWorksheet.Cells(1, "C") = BOR!MBMGroupCompany
                    objWorksheet.Cells(2, "C") = BOR!MBMPayorEffDate & " To " & BOR!MBMPayorEXPDate
                    objWorksheet.Cells(3, "C") = BOR!BordxDate
                    objWorksheet.Cells(4, "C") = TxtBordxNo
                    objWorksheet.Cells(5, "C") = BOR!BordxDate
                
                End If
                tmpSeqNo = CDbl(tmpSeqNo) + 1
                objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = BOR!MBMDeptCode
                objWorksheet.Cells(7 + CDbl(PVCnt), "C") = BOR!MBMDepartment
                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = BOR!MCDays
                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = BOR!CLMTotalActual
                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = 0
                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = BOR!CLMBordxAmt
                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = "GP"
                tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
            BOR.MoveNext
        Loop
        BOR.Close
        Set BOR = Nothing
        'Specialist
        sqlbor = ""
        sqlbor = "Select * from VIEW_CISMBMInquiry_SPSummary where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
        BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not BOR.EOF
            PVCnt = PVCnt + 1
                If PVCnt = "1" Then
                    
                    objWorksheet.Cells(1, "C") = BOR!MBMGroupCompany
                    objWorksheet.Cells(2, "C") = BOR!MBMPayorEffDate & " To " & BOR!MBMPayorEXPDate
                    objWorksheet.Cells(3, "C") = BOR!BordxDate
                    objWorksheet.Cells(4, "C") = TxtBordxNo
                    objWorksheet.Cells(5, "C") = BOR!BordxDate
                End If
                tmpSeqNo = CDbl(tmpSeqNo) + 1
                objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = BOR!MBMDeptCode
                objWorksheet.Cells(7 + CDbl(PVCnt), "C") = BOR!MBMDepartment
                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = BOR!MCDays
                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = Format(BOR!CLMTotalActual, "#,##0.00")
                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = 0
                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = Format(BOR!CLMBordxAmt, "#,##0.00")
                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = "SP"
                tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
            BOR.MoveNext
        
        Loop
        PVCnt = PVCnt + 2
        objWorksheet.Cells(7 + CDbl(PVCnt), "E") = Format(tmpTotalActual, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "F") = Format(0, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "G") = Format(tmpTotalBordxAmt, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "E").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "F").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "G").Font.Bold = True
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        If Err.Description <> "" Then
            Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
        End If
        Exit Sub
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub












