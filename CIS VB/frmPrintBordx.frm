VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintBordx 
   Caption         =   "Print Bordx"
   ClientHeight    =   8208
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8208
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Print Bordx"
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
         TabIndex        =   12
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   11775
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   4695
      End
      Begin VB.ComboBox CboYear 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrintBordx.frx":0000
         Left            =   9840
         List            =   "frmPrintBordx.frx":0002
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox CboMonth 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrintBordx.frx":0004
         Left            =   7200
         List            =   "frmPrintBordx.frx":002C
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Left            =   9000
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Month"
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   6135
      Left            =   0
      TabIndex        =   16
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
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   7200
      Width           =   11775
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   480
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3120
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   495
         Left            =   9480
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   495
         Left            =   8760
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   10200
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   495
         Left            =   10920
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintSummary 
         Caption         =   "&Print Summary"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5760
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintDetails 
         Caption         =   "Print &Detail"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSummaryPreview 
         Caption         =   "S&ummary Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDetailPreview 
         Caption         =   "De&tail Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7680
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtPStatus 
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin Crystal.CrystalReport crptSummary 
         Left            =   1080
         Top             =   720
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_PrintBordx 
         Left            =   1560
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPrintBordx"
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
Dim Prm2 As New ADODB.Parameter
Dim TempMBMBranch As String
Dim CLM As New ADODB.Recordset
Dim sql As String
Dim sql2 As String
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
    
    bPrintPass = False
    CancelPrinting = False
    Screen.MousePointer = vbHourglass
    tempPath = "\\PLUTO\HISSaturn\Conn\MediConnCISWSPluto.UDL"
    MediConnCISWSPluto.Open "File Name=" & tempPath
    
    strsqlPlutoBor = "Select BordxRefNo,BordxDate from TempBordxSummary where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
    PlutoBor.Open strsqlPlutoBor, MediConnCISWSPluto, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PlutoBor.EOF
        PLUTOStatus = True
        tmpbordxdate = PlutoBor!BordxDate
        PlutoBor.MoveNext
    Loop
    PlutoBor.Close
    Set PlutoBor = Nothing

    If PLUTOStatus = False Then
    
        sqlbor = " And BordxRefNo = '" & ChkStr(Trim(TxtBordxNo)) & "'"
        
        Set cmd.ActiveConnection = MediConnCISWS
        cmd.CommandText = "ClaimBordxModule"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sqlbor)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "12")
        cmd.Parameters.Append Prm2
        BOR.CursorLocation = adUseClient
        BOR.CursorType = adOpenForwardOnly
        BOR.CacheSize = 100
        Set BOR = cmd.Execute
        
        'Set cmd2.ActiveConnection = MediConnCISWS
        'cmd2.CommandText = "ClaimBordxModule"
        'cmd2.CommandType = adCmdStoredProc
        'cmd2.CommandTimeout = 300
        'Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sqlbor)
        'cmd2.Parameters.Append Prm3
        'Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "20")
        'cmd2.Parameters.Append Prm4
        'CLM.CursorLocation = adUseClient
        'CLM.CursorType = adOpenForwardOnly
        'CLM.CacheSize = 100
        'Set CLM = cmd2.Execute
    
       ' If CLM.EOF = False Then
       '     'lstPrintBordx.ListItems.Clear
       '     'TotalAmt = 0
       '     Do While Not CLM.EOF
            
       '         sqlbor = ""
       '         sqlbor = " Select * From View_CIS_PrintBordx where CasNumber = '" & Trim(CLM!CasNumber) & "'"
       '          BOR.Open sqlbor, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
       '        ' sqlbor = "Select * from VIEW_CISClaimBordx_NS where CASNumber = '" & Trim(CLM!CasNumber) & "'"
       '        ' BOR.Open sqlbor, MediConnCISWS, adOpenForwardOnly, adLockReadOnly

                Do While Not BOR.EOF
                    
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
                    
                    sqlinsert = ""
                    sqlinsert = "Insert into TempBordxSummary(CASNumber, BordxRefNo, CLMTotalActual, CLMCoPayColl," & _
                                                                "CLMBordxAmt, CLMFinalApprove,CLMMCoProcFee,MBMnumber," & _
                                                                "CLMPayableByMedix2C,CLMPayableByMedix2M, CLMTotalApproved, " & _
                                                                "DebitCreditRefNo, CLMWSRemarks,PatientName,Relation, " & _
                                                                "CoverID,DateVisit,MBMName,GRPCompany,MBMPolicyNo,PLNCode, " & _
                                                                "MBMEmployeeNo,HLTCode,MBMCName,MBMMemberType,PAYCode,INSCode, " & _
                                                                "CLNName,CLNCode,MBMCRELCode,RELCode,MBMGroupCompany,MBMDepartment, " & _
                                                                "Remarks,BordxPrintedStatus,BordxDate,BordxPrintMonth,BordxPrintYear, " & _
                                                                "ReasonDesc1,ReasonDesc2,ReasonDesc3,CASDiagnosis5,MBMCostCenter,GRPBordxFormat, MBMBranch) " & _
                                "Values('" & BOR!CasNumber & "', '" & BOR!Bordxrefno & "', " & BOR!CLMTotalActual & ", " & BOR!CLMCoPayColl & "," & _
                                                                "" & BOR!CLMBordxAmt & ", " & BOR!CLMFinalApprove & ", " & BOR!CLMMCOProcFee & ", '" & BOR!MBMNumber & "', " & _
                                                                "" & IIf(IsNumeric(BOR!CLMPayableByMedix2C), BOR!CLMPayableByMedix2C, "0") & ", " & IIf(IsNumeric(BOR!CLMPayableByMedix2M), BOR!CLMPayableByMedix2M, "0") & ", " & IIf(IsNumeric(BOR!CLMTotalApproved), BOR!CLMTotalApproved, "0") & ", " & _
                                                                "'" & BOR!DebitCreditRefNo & "', '" & BOR!CLMWSRemarks & "', '" & TempPatientName & "', '" & BOR!Relation & "', " & _
                                                                "'" & BOR!coverid & "','" & tmpDateVisit & "','" & TempMBMName & "','" & BOR!GRPCompany & "','" & BOR!MBMPolicyNo & "','" & BOR!PLNCode & "', " & _
                                                                "'" & BOR!MBMEmployeeNo & "','N','" & TempMBMCName & "','" & BOR!MBMMemberType & "','" & BOR!PAYCode & "','" & BOR!INSCode & "', " & _
                                                                "'" & TempCLNName & "', '" & BOR!CLNCode & "', '" & BOR!MBMCRELCode & "','" & BOR!RELCODE & "', '" & tempGrpCompany & "','" & TempMBMDept & "', " & _
                                                                "'" & TempRemarks & "','" & IIf(Trim(BOR!BordxPrintedStatus) = "TRUE", "1", "0") & "','" & tmpbordxdate & "', '" & BOR!BORDXPrintMonth & "','" & BOR!BORDXPrintYear & "', " & _
                                                                "'" & BOR!ReasonDesc1 & "','" & BOR!ReasonDesc2 & "','" & BOR!ReasonDesc3 & "','" & TempCASDiagnosis5 & "', '" & BOR!MBMCostCenter & "', '" & tmpRptFormat & "', '" & TempMBMBranch & "' )"
                    MediConnCISWSPluto.Execute sqlinsert
                               'BordxMNRBSql = "Insert Into SUNPosting(TransRef, BordxMNRBStatus," & _
                                '            "PostStatus, PostType, PostDate, PostBy, " & _
                                '            "AccMonth, AccYear, BatchNo) " & _
                                '            "Values('" & TempReceiptNo & "', '1','1','" & PostingType & _
                                '            "','" & Format(Date, "yyyy/mm/dd") & _
                                '            "','" & Logon & "','" & Mid(cboAccPeriod, 1, 2) & _
                                '            "','" & Mid(cboAccPeriod, 4, 4) & "','" & BatchNo & "')"
                                'medixmasterconnWS.Execute BordxMNRBSql
            
                
                'BOR.MoveNext
                'Loop
                'End If
                'BOR.Close
                'Set BOR = Nothing
                
                BOR.MoveNext
                Loop
            'End If
            BOR.Close
            Set BOR = Nothing
            End If
            MediConnCISWSPluto.Close
            Set MediConnCISWSPluto = Nothing
    
    crptSummary.SelectionFormula = "{VIEW_CISClaimBordx_NS.BordxRefNo} = '" & Trim(TxtBordxNo) & "'"
    
    If tmpRptFormat = "" Then
        If tmpRptByDept = False Then
            If GetFFS(GRPCompany) = True Then
                crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail_FFS.rpt"
            Else
                If tmpRptType = "M" Then
                    crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail_Medix.rpt"
                Else
                    crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail.rpt"
                End If
            End If
            crptSummary.SortFields(0) = "+{VIEW_CISClaimBordx_NS.CLNName}"
            crptSummary.SortFields(1) = "+{VIEW_CISClaimBordx_NS.DateVisit}"
        Else
            If GetFFS(GRPCompany) = True Then
                crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail_Dept_FFS.rpt"
            Else
            'shan 121010
                'crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail_Dept.rpt"
                If tmpRptType = "M" Then
                    crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail_Dept_medix.rpt"
                Else
                    If tmpRptThphoidType = "Y" Then
                        crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail_Dept_Typhoid.rpt"
                    Else
                        crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail_DeptNew.rpt"
                    End If
                End If
            End If
            crptSummary.SortFields(0) = "+{VIEW_CISClaimBordx_NS.MBMDepartment}"
            'crptSummary.SortFields(1) = "+{VIEW_CISClaimBordx_NS.CLNName}"
            crptSummary.SortFields(1) = "+{VIEW_CISClaimBordx_NS.MBMName}"
        End If
    ElseIf tmpRptFormat = "1" Then
            crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Detail_Dept_MDec.rpt"
            crptSummary.SortFields(0) = "+{VIEW_CISClaimBordx_NS.MBMDepartment}"
            crptSummary.SortFields(1) = "+{VIEW_CISClaimBordx_NS.CLNName}"

    End If
    
    If cmdPrintDetails Then
        crptSummary.Destination = crptToPrinter
        bPrintPass = PrintFunction
        If bPrintPass = False Or CancelPrinting = True Then
            Screen.MousePointer = Default
            Exit Sub
        Else
            If tmpPrintStatus = 0 Then
                Call UpdateBordx
            End If
            crptSummary.PrintReport
        End If
    Else
        crptSummary.Destination = crptToWindow
        crptSummary.WindowState = crptMaximized
        crptSummary.PrintReport
    End If
    Screen.MousePointer = Default
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Function PrintFunction() As Boolean
On Error GoTo cancelPrint
        If crptSummary Is Nothing Then Exit Function
        With dlgCommonDialog_PrintBordx
            .DialogTitle = "Print Crystal Report "
            .CancelError = True
            .Flags = cdlPDNoSelection + cdlPDNoPageNums
            .ShowPrinter
            If Err <> MSComDlg.cdlCancel Then
            End If
        End With
        PrintFunction = True
        Exit Function
cancelPrint:
        PrintFunction = False
        CancelPrinting = True
End Function

Private Sub SummaryPrint()

Const METHOD_NAME = "SummaryPrint"
On Error GoTo debugs

Dim StrSql1 As String
Dim tempPath As String
Dim tempGrpCompany As String
Dim TempMBMDept As String
Dim strsqlPlutoBor As String
Dim sqlbor As String
Dim PlutoBor As New ADODB.Recordset
Dim BOR As New ADODB.Recordset
Dim PLUTOStatus As Boolean
Dim tmpbordxdate As String
Dim sqlinsert As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter

    bPrintPass = False
    CancelPrinting = False
    Screen.MousePointer = vbHourglass
    
    tempPath = "\\PLUTO\HISSaturn\Conn\MediConnCISWSPluto.UDL"
    MediConnCISWSPluto.Open "File Name=" & tempPath
    
    strsqlPlutoBor = "Select BordxRefNo,BordxDate from TempBordxSum where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
    PlutoBor.Open strsqlPlutoBor, MediConnCISWSPluto, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PlutoBor.EOF
        PLUTOStatus = True
        tmpbordxdate = PlutoBor!BordxDate
        PlutoBor.MoveNext
    Loop
    PlutoBor.Close
    Set PlutoBor = Nothing

    If PLUTOStatus = False Then
       ' sqlbor = "Select * from VIEW_CISClaimBordxNS_Below_DEPT where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
       ' BOR.Open sqlbor, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
        sqlbor = " And BordxRefNo = '" & ChkStr(Trim(TxtBordxNo)) & "'"
        
        Set cmd.ActiveConnection = MediConnCISWS
        cmd.CommandText = "ClaimBordxModule"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sqlbor)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "19")
        cmd.Parameters.Append Prm2
        BOR.CursorLocation = adUseClient
        BOR.CursorType = adOpenForwardOnly
        BOR.CacheSize = 100
        Set BOR = cmd.Execute

    
        Do While Not BOR.EOF
            
            tmpbordxdate = Format(BOR!BordxDate, "yyyy/mm/dd")
            tempGrpCompany = Replace(Trim(BOR!GroupCompany), "'", "''")
            TempMBMDept = Replace(Trim(BOR!MBMDepartment), "'", "''")

            sqlinsert = ""
            sqlinsert = "Insert into TempBordxSum(BordxRefNo,TotalClaim,TotalAmount, " & _
                                                "GroupCompany,BORDXPrintMonth,BORDXPrintYear,MBMDepartment," & _
                                                "GRPCoAdd1, GRPCoAdd2,GRPCoAdd3,CLMTotalActual,CLMCoPayColl," & _
                                                "CLMFinalApprove,CLMMCOProcFee,BordxDate,MBMBranch) " & _
                        "Values('" & BOR!Bordxrefno & "', " & IIf(IsNumeric(BOR!TotalClaim), BOR!TotalClaim, "0") & ", " & IIf(IsNumeric(BOR!TotalAmount), BOR!TotalAmount, "0") & ", " & _
                                "'" & BOR!GroupCompany & "', '" & BOR!BORDXPrintMonth & "', '" & BOR!BORDXPrintYear & "', '" & BOR!MBMDepartment & "', " & _
                                "'" & BOR!GRPCoAdd1 & "', '" & BOR!GRPCoAdd2 & "', '" & BOR!GRPCoAdd3 & "', " & _
                                "" & IIf(IsNumeric(BOR!CLMTotalActual), BOR!CLMTotalActual, "0") & ", " & IIf(IsNumeric(BOR!CLMCoPayColl), BOR!CLMCoPayColl, "0") & ", " & _
                                "" & IIf(IsNumeric(BOR!CLMFinalApprove), BOR!CLMFinalApprove, "0") & ", " & IIf(IsNumeric(BOR!CLMMCOProcFee), BOR!CLMMCOProcFee, "0") & ", '" & tmpbordxdate & "', '" & IIf(Len(BOR!MBMBranch), BOR!MBMBranch, "") & "')"
                   MediConnCISWSPluto.Execute sqlinsert
                       'BordxMNRBSql = "Insert Into SUNPosting(TransRef, BordxMNRBStatus," & _
                        '            "PostStatus, PostType, PostDate, PostBy, " & _
                        '            "AccMonth, AccYear, BatchNo) " & _
                        '            "Values('" & TempReceiptNo & "', '1','1','" & PostingType & _
                        '            "','" & Format(Date, "yyyy/mm/dd") & _
                        '            "','" & Logon & "','" & Mid(cboAccPeriod, 1, 2) & _
                        '            "','" & Mid(cboAccPeriod, 4, 4) & "','" & BatchNo & "')"
                        'medixmasterconnWS.Execute BordxMNRBSql
    
        
        BOR.MoveNext
        Loop
    End If
    MediConnCISWSPluto.Close
    Set MediConnCISWSPluto = Nothing
    
    crptSummary.SelectionFormula = "{VIEW_CISClaimBordxNS_Below_DEPT.BordxRefNo} = '" & Trim(TxtBordxNo) & "'"

    If tmpRptByDept = False Then
        If GetFFS(GRPCompany) = True Then
            crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary_FFS.rpt"
        Else
            If tmpRptType = "M" Then
                crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary_Medix.rpt"
            Else
                crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary.rpt"
            End If
        End If
    Else
        If GetFFS(GRPCompany) = True Then
            crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary_Dept_FFS.rpt"
        Else
        'shan 121010
            'crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary_Dept.rpt"
            If tmpRptType = "M" Then
                crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary_Dept_Medix.rpt"
            Else
                crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary_DeptNew.rpt"
            End If
        End If
        crptSummary.SelectionFormula = "{VIEW_CISClaimBordxNS_Below_DEPT.BordxRefNo} = '" & Trim(TxtBordxNo) & "'"
        crptSummary.SortFields(0) = "+{VIEW_CISClaimBordxNS_Below_DEPT.MBMDepartment}"
    End If
    If CmdPrintSummary Then
        crptSummary.Destination = crptToPrinter
        bPrintPass = PrintFunction
        If bPrintPass = False Or CancelPrinting = True Then
            Screen.MousePointer = Default
            Exit Sub
        Else
            If tmpPrintStatus = 0 Then
                Call UpdateBordx
            End If
            crptSummary.PrintReport
        End If
    Else
        crptSummary.Destination = crptToWindow
        crptSummary.WindowState = crptMaximized
        crptSummary.PrintReport
    End If
    Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdDetailPreview_Click()
        If VerifyPrint = True Then
            If tmpRptFormat = "2" Then
                Call PrintTemplate
            Else
                Call DetailPrint
            End If
        End If
    crptSummary.Reset
End Sub

Private Sub cmdPrintDetails_Click()
            If tmpRptFormat = "2" Then
                Call PrintTemplate
            Else
                Call DetailPrint
            End If
        crptSummary.Reset
End Sub

Private Sub CmdPrintSummary_Click()
    If VerifyPrint = True Then
        Call SummaryPrint
    End If
End Sub

Private Sub CmdSummaryPreview_Click()
    If VerifyPrint = True Then
        Call SummaryPrint
    End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPrintBordx")
Dim AA As Integer
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    CboYear.Clear
    For AA = 2002 To Year(Date) + 2
        CboYear.AddItem AA
    Next AA
    CmdPrintSummary.Enabled = False
    cmdPrintDetails.Enabled = False
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

Private Sub CmdSearch_Click()
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
            strAppend = " And BordxRefNo = '" & ChkStr(Trim(TxtBordxNo)) & "'"
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

Private Sub cmdExit_Click()
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
Dim Prm2 As New ADODB.Parameter
Dim Check As Boolean
Dim CLM As New ADODB.Recordset
Dim sql As String
Dim sql2 As String
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim strAppend2 As String
    tmpRptType = ""
    tmpRptThphoidType = ""
    TotRec = 0
    CurrRec = 0
    Check = False
    
    'sql = " Select  CASNumber, BordxRefNo From Claim where BordxRefNo = '" & TxtBordxNo & "'"
    'CLM.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        Set cmd2.ActiveConnection = MediConnCISWS
        cmd2.CommandText = "ClaimBordxModule"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd2.Parameters.Append Prm3
        Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "12")
        cmd2.Parameters.Append Prm4
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
                        TotRec = CLM.RecordCount
                        lblTotal = TotRec
                         With CLM
                            GRPCompany = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
                            Set Record = lstPrintBordx.ListItems.Add(, , !PLNCode)
                            Record.SubItems(1) = IIf(Len(Trim(!CasNumber)), Trim(!CasNumber), "")
                            Record.SubItems(2) = IIf(Len(Trim(!PatientName)), Trim(!PatientName), "")
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
      '  MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\MediConnCISPayment.UDL"
      '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            
            Call GetBordxInfo
            
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
       ' MediConnCISWS.Close
       ' Set MediConnCISWS = Nothing
       ' MediConnCISPayment.Close
       ' Set MediConnCISPayment = Nothing
        
    End If
End Sub

Private Sub UpdateBordx()
Const METHOD_NAME = "UpdateBordx"
On Error GoTo debugs

Dim strsql As String
    strsql = " update Bordx set " & _
             " BordxPrintedDate = '" & Format(Date, "YYYY/MM/DD") & "', " & _
             " BORDXPrintMonth = '" & ChkStr(Trim(CboMonth)) & "', " & _
             " BORDXPrintYear = '" & Trim(CboYear) & "', " & _
             " BordxPrintedStatus = '1', " & _
             " BordxPrintedBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
    MediConnCISWS.Execute strsql
    
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
End Function

Private Sub PrintTemplate()
Const METHOD_NAME = "PrintTemplate"
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
Dim tmpTotalVisits As String
Dim tmpTotalMC As String
        
        tmpTotalBordxAmt = 0
        PVCnt = 0
        tmpSeqNo = 0
        tmpTotalActual = 0
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Airod OP Template.xls")
            'Set objWorkbook = objExcel.Workbooks.Add("C:\Template\Airod OP Template.xls")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If

        'GP
        'sqlbor = "Select * from VIEW_CISMBMInquiry_GPCASE where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
        sqlbor = "Select * from VIEW_CISMBMInquiry_GPCASE_New where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
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
                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = BOR!CasNumber
                objWorksheet.Cells(7 + CDbl(PVCnt), "C") = BOR!MBMEmployeeNo
                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = BOR!MBMName
                If BOR!coverid = "00" Then
                    objWorksheet.Cells(7 + CDbl(PVCnt), "E") = BOR!MBMName
                Else
                    objWorksheet.Cells(7 + CDbl(PVCnt), "E") = BOR!MBMCName
                End If
                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = BOR!MBMDeptCode
                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = BOR!MBMDepartment
                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = BOR!MCDays
                objWorksheet.Cells(7 + CDbl(PVCnt), "I") = BOR!DateVisit
                
                
                If BOR!INVPayTo = "C" Then
                    objWorksheet.Cells(7 + CDbl(PVCnt), "J") = BOR!CLNName
                Else
                    objWorksheet.Cells(7 + CDbl(PVCnt), "J") = BOR!NPClinic
                End If
                'objWorksheet.Cells(7 + CDbl(PVCnt), "K") = BOR!DiagnosisAllDesc & " " & BOR!CASDiagnosis5
                objWorksheet.Cells(7 + CDbl(PVCnt), "K") = UCase(BOR!treatment)
                objWorksheet.Cells(7 + CDbl(PVCnt), "L") = BOR!CLMTotalActual
                objWorksheet.Cells(7 + CDbl(PVCnt), "M") = 0
                objWorksheet.Cells(7 + CDbl(PVCnt), "N") = BOR!CLMBordxAmt
                objWorksheet.Cells(7 + CDbl(PVCnt), "O") = "GP"
                tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
                
                'objWorksheet.Cells(7 + CDbl(PVCnt), "J").Borders.LineStyle = True
                'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                
            BOR.MoveNext
        
        
        'BOR.MoveNext
        Loop
        BOR.Close
        Set BOR = Nothing
        'Specialist
        sqlbor = ""
        sqlbor = "Select * from VIEW_CISMBMInquiry_Spe_CASE_New where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
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
                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = BOR!CasNumber
                objWorksheet.Cells(7 + CDbl(PVCnt), "C") = BOR!MBMEmployeeNo
                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = BOR!MBMName
                If BOR!coverid = "00" Then
                    objWorksheet.Cells(7 + CDbl(PVCnt), "E") = BOR!MBMName
                Else
                    objWorksheet.Cells(7 + CDbl(PVCnt), "E") = BOR!MBMCName
                End If
                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = BOR!MBMDeptCode
                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = BOR!MBMDepartment
                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = BOR!MCDays
                objWorksheet.Cells(7 + CDbl(PVCnt), "I") = BOR!DateVisit
                
                
                If BOR!INVPayTo = "C" Then
                    objWorksheet.Cells(7 + CDbl(PVCnt), "J") = BOR!CLNName
                Else
                    objWorksheet.Cells(7 + CDbl(PVCnt), "J") = BOR!NPClinic
                End If
                'objWorksheet.Cells(7 + CDbl(PVCnt), "K") = BOR!DiagnosisAllDesc & " " & BOR!CASDiagnosis5
                objWorksheet.Cells(7 + CDbl(PVCnt), "K") = UCase(BOR!treatment)
                objWorksheet.Cells(7 + CDbl(PVCnt), "L") = Format(BOR!CLMTotalActual, "#,##0.00")
                objWorksheet.Cells(7 + CDbl(PVCnt), "M") = 0
                objWorksheet.Cells(7 + CDbl(PVCnt), "N") = Format(BOR!CLMBordxAmt, "#,##0.00")
                objWorksheet.Cells(7 + CDbl(PVCnt), "O") = "SP"
                tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
                
                
            BOR.MoveNext
        
        
        'BOR.MoveNext
        Loop
        PVCnt = PVCnt + 2
        objWorksheet.Cells(7 + CDbl(PVCnt), "L") = Format(tmpTotalActual, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "M") = Format(0, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "N") = Format(tmpTotalBordxAmt, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "L").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "M").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "N").Font.Bold = True
        'objWorksheet.Cells(7 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        PVCnt = PVCnt + 1
        objWorksheet.Cells(7 + CDbl(PVCnt), "B") = "Kindly note that payment term is 14 days and payment should be issued to Health Connect Sdn Bhd"
        objWorksheet.Cells(7 + CDbl(PVCnt), "B").Font.Bold = True
        
        BOR.Close
        Set BOR = Nothing
        
        tmpTotalBordxAmt = 0
        PVCnt = 0
        tmpSeqNo = 0
        tmpTotalActual = 0
        tmpTotalVisits = 0
        tmpTotalMC = 0

        
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
                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = BOR!TotalVisits
                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = IIf(IsNumeric(BOR!MCDays), BOR!MCDays, 0)
                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = BOR!CLMTotalActual
                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = 0
                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = BOR!CLMBordxAmt
                objWorksheet.Cells(7 + CDbl(PVCnt), "I") = "GP"
                tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
                tmpTotalVisits = CDbl(tmpTotalVisits) + CDbl(BOR!TotalVisits)
                tmpTotalMC = CDbl(tmpTotalMC) + CDbl(BOR!MCDays)
                'objWorksheet.Cells(7 + CDbl(PVCnt), "J").Borders.LineStyle = True
                'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                
            BOR.MoveNext
        
        
        'BOR.MoveNext
        Loop
        BOR.Close
        Set BOR = Nothing
        'Specialist
        sqlbor = ""
        sqlbor = "Select * from VIEW_CISMBMInquiry_SPSummary where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
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
                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = BOR!TotalVisits
                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = IIf(IsNumeric(BOR!MCDays), BOR!MCDays, 0)
                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = BOR!CLMTotalActual
                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = 0
                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = BOR!CLMBordxAmt
                objWorksheet.Cells(7 + CDbl(PVCnt), "I") = "SP"
                tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
                tmpTotalVisits = CDbl(tmpTotalVisits) + CDbl(BOR!TotalVisits)
                tmpTotalMC = CDbl(tmpTotalMC) + CDbl(BOR!MCDays)
                
                
            BOR.MoveNext
        
        
        'BOR.MoveNext
        Loop
        PVCnt = PVCnt + 2
        
        
        objWorksheet.Cells(7 + CDbl(PVCnt), "D") = tmpTotalVisits
        objWorksheet.Cells(7 + CDbl(PVCnt), "E") = tmpTotalMC
        objWorksheet.Cells(7 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "G") = Format(0, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "H") = Format(tmpTotalBordxAmt, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "D").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "E").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "F").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "G").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "H").Font.Bold = True
        'objWorksheet.Cells(7 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlContinuous
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
        

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
Exit Sub

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
                
                'objWorksheet.Cells(7 + CDbl(PVCnt), "J").Borders.LineStyle = True
                'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                
            BOR.MoveNext
        
        
        'BOR.MoveNext
        Loop
        BOR.Close
        Set BOR = Nothing
        'Specialist
        sqlbor = ""
        sqlbor = "Select * from VIEW_CISMBMInquiry_SPSummary where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
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
                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = Format(BOR!CLMTotalActual, "#,##0.00")
                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = 0
                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = Format(BOR!CLMBordxAmt, "#,##0.00")
                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = "SP"
                tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
                
                
            BOR.MoveNext
        
        
        'BOR.MoveNext
        Loop
        PVCnt = PVCnt + 2
        objWorksheet.Cells(7 + CDbl(PVCnt), "E") = Format(tmpTotalActual, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "F") = Format(0, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "G") = Format(tmpTotalBordxAmt, "#,##0.00")
        objWorksheet.Cells(7 + CDbl(PVCnt), "E").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "F").Font.Bold = True
        objWorksheet.Cells(7 + CDbl(PVCnt), "G").Font.Bold = True
        'objWorksheet.Cells(7 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        
        
        
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




