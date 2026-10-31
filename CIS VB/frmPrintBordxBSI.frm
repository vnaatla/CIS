VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintBordxBSI 
   Caption         =   "Print Bordx (Berjaya Sompo)"
   ClientHeight    =   8676
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11868
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8676
   ScaleWidth      =   11868
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Print Bordx (Berjaya Sompo)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   0
      TabIndex        =   13
      Top             =   7200
      Width           =   11775
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3240
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPStatus 
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdDetailPreview 
         Caption         =   "De&tail Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5280
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdSummaryPreview 
         Caption         =   "S&ummary Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6120
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintDetails 
         Caption         =   "Print &Detail"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7920
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdPrintSummary 
         Caption         =   "&Print Summary"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6960
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   495
         Left            =   10920
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   10200
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   495
         Left            =   8760
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   495
         Left            =   9480
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11775
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   3720
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   200
         Width           =   1575
      End
      Begin VB.ComboBox CboYear 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrintBordxBSI.frx":0000
         Left            =   5760
         List            =   "frmPrintBordxBSI.frx":0002
         TabIndex        =   4
         Top             =   195
         Width           =   1695
      End
      Begin VB.ComboBox CboMonth 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrintBordxBSI.frx":0004
         Left            =   3480
         List            =   "frmPrintBordxBSI.frx":002C
         TabIndex        =   3
         Top             =   195
         Width           =   1695
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         ItemData        =   "frmPrintBordxBSI.frx":0092
         Left            =   1800
         List            =   "frmPrintBordxBSI.frx":0094
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ComboBox CboReportType 
         Height          =   315
         ItemData        =   "frmPrintBordxBSI.frx":0096
         Left            =   7680
         List            =   "frmPrintBordxBSI.frx":00A9
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   675
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "Company"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblDept 
         Caption         =   "Department"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Month"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Report Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   6135
      Left            =   0
      TabIndex        =   28
      Top             =   1080
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   10816
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
Attribute VB_Name = "frmPrintBordxBSI"
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
    If Trim(CboMonth) = "" Or Trim(CboYear) = "" Then
        MsgBox "Please select bordx month or bordx year", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Call PrintTemplate
    End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPrintBordxBSI")
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
End Sub

'******************** Start Command Button ********************************
Private Sub cmdClear_Click()
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

Private Sub cmdStop_Click()
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
        Call GetBordxInfo
    End If
End Sub

Private Sub UpdateBordx()
Const METHOD_NAME = "UpdateBordx"
On Error GoTo debugs

Dim StrSql As String
    StrSql = " update Bordx set " & _
             " BordxPrintedDate = '" & Format(Date, "YYYY/MM/DD") & "', " & _
             " BORDXPrintMonth = '" & ChkStr(Trim(CboMonth)) & "', " & _
             " BORDXPrintYear = '" & Trim(CboYear) & "', " & _
             " BordxPrintedStatus = '1', " & _
             " BordxPrintedBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
    MediConnCISWS.Execute StrSql
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub GetBordxInfo()
Const METHOD_NAME = "GetBordxInfo"
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

Private Sub PrintTemplate()
Const METHOD_NAME = "PrintTemplate"
On Error GoTo debugs

Dim StrSql1 As String
Dim StrSql2 As String
Dim Subsidiary As New ADODB.Recordset
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
Dim tmpTotalVisits As String
Dim tmpTotalMC As String
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim strsqlCOV As String
Dim RstCov As New ADODB.Recordset
Dim tmpPatientName As String
Dim tmpSubMBMPaid As Double
Dim totalMBMPaid As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim PLN As New ADODB.Recordset
Dim strsqlPLN As String
Dim PlnDesc As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim StrSql As String
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
Dim tmpTotalCases As String
Dim lblDate As String
Dim RecFound As Boolean
Dim sValue As String
Dim SubTotal As Double
Dim IsFirst As Boolean
Dim Rcount As Integer
Dim Diag As String
Dim SubSid As String
Dim tmpBordxNo As String
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim tmpMBMGroupCompany As String
SubSid = ""
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
            
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\BSIOPBordx.xlt")
            
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
        
                    StrSql = ""
                    StrSql = " BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd4.ActiveConnection = MediConnCISWS
                    cmd4.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd4.CommandType = adCmdStoredProc
                    cmd4.CommandTimeout = 300
                    Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd4.Parameters.Append Prm7
                    Set prm8 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 9)
                    cmd4.Parameters.Append prm8
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd4.Execute
                    
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
                    
                    Do While Not rst2.EOF
                        RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(1)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(1)
                                               
                            objWorksheet.Cells(10, "K") = rst2!Bordxrefno
                            objWorksheet.Cells(12, "K") = "30th" & " " & CboMonth & " " & CboYear
                            objWorksheet.Cells(20, "B") = "Outpatient medical cost incurred for the month of" & " " & CboMonth & " " & CboYear
                       
                            objWorksheet.Cells(20, "L") = rst2!TotalCases
                            objWorksheet.Cells(20, "M") = Format(rst2!TotalActual, "#,##0.00")
                            tmpTotalActual = CDbl(rst2!TotalActual)
                            tmpBordxNo = rst2!Bordxrefno
                            tmpMBMGroupCompany = rst2!MBMGroupCompany
                        rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    
                If RecFound = True Then
                    
                    Set objword = Nothing
                    If objword Is Nothing Then
                        Set objword = New Word.Application
                        objword.Visible = True
                        'Set objdoc = objword.Documents.Add(crdPath & "CISTemplate\CoverNoteBSI1.dot")
                        Set objdoc = objword.Documents.Add(LocCardPath & "Claim\CoverNoteBSI.dot")
                        
                    End If
                                        
                    If tmpBordxNo <> "" Then
                        objdoc.Bookmarks("BordxNo").Range.Text = tmpBordxNo
                    End If
                    
                    If tmpTotalActual <> "" Then
                        objdoc.Bookmarks("BordxAmt").Range.Text = Format(tmpTotalActual, "#,##0.00")
                    End If
                                    
                    objdoc.Bookmarks("BillingDate").Range.Text = "30th" & " " & CboMonth & " " & CboYear
                    
                    objdoc.Bookmarks("Company").Range.Text = tmpMBMGroupCompany
                    
                    StrSql = ""
                    StrSql = " CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 300
                    Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, StrSql)
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
                        Set objWorksheet = objWorkbook.Sheets(2)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        If PVCnt = 1 Then
                            objWorksheet.Cells(2, "K") = rst2!MBMGroupCompany
                            objWorksheet.Cells(3, "K") = TxtBordxNo
                            objWorksheet.Cells(4, "K") = "30th" & " " & CboMonth & " " & CboYear
                            'objWorksheet.Cells(6, "L") = IIf(Len(SubSid), SubSid, "")
                        End If
                            Rcount = Rcount + 1
                            objWorksheet.Cells(14 + CDbl(PVCnt), "A") = Rcount
                            objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CasNumber
                            objWorksheet.Cells(14 + CDbl(PVCnt), "C") = Right(PlnDesc, 2)
                            If rst2!coverid = "00" Then
                                objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "E"
                            Else
                                If rst2!MBMCRelcode = "H" Then
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "S"
                                ElseIf rst2!MBMCRelcode = "W" Then
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "S"
                                ElseIf rst2!MBMCRelcode = "S" Then
                                    objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "C"
                                ElseIf rst2!MBMCRelcode = "D" Then
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
                      '      tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMTotalActual)
                      ' IsFirst = False
                       rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    RecFound = False
                    
                       '  objWorksheet.Cells(14 + CDbl(PVCnt + 1), "U").Borders(xlEdgeTop).LineStyle = xlContinuous
                       '  objWorksheet.Cells(14 + CDbl(PVCnt + 1), "U").Borders(xlEdgeBottom).LineStyle = xlDouble
                       '  objWorksheet.Cells(14 + CDbl(PVCnt + 1), "U").Font.Bold = True
                       '  objWorksheet.Cells(14 + CDbl(PVCnt + 1), "U") = "Sub Total"
                       '  objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                       '  objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                       '  objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V").Font.Bold = True
                       '  objWorksheet.Cells(14 + CDbl(PVCnt + 1), "V") = Format(SubTotal, "#,##0.00")
                    PVCnt = PVCnt + 2
                    objWorksheet.Cells(14 + CDbl(PVCnt), "U").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "U").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "U").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "U") = "Total"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(tmpsubAct, "#,##0.00")
            End If
        
        
        
        'End If
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














