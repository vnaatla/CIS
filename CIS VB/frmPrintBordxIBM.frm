VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintBordxIBM 
   Caption         =   "Print Bordx - SEB/CIMB"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   7200
      Width           =   11775
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3120
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   495
         Left            =   9480
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   495
         Left            =   8760
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   10200
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   495
         Left            =   10920
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintSummary 
         Caption         =   "&Print Summary"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintDetails 
         Caption         =   "Print &Detail"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7920
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSummaryPreview 
         Caption         =   "S&ummary Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4920
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdDetailPreview 
         Caption         =   "De&tail Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4560
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtPStatus 
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin Crystal.CrystalReport crptSummary 
         Left            =   1080
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
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
         TabIndex        =   25
         Top             =   360
         Width           =   735
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
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   11775
      Begin VB.ComboBox CmbGrpcomapny 
         Height          =   315
         ItemData        =   "frmPrintBordxIBM.frx":0000
         Left            =   7080
         List            =   "frmPrintBordxIBM.frx":0002
         TabIndex        =   28
         Top             =   120
         Width           =   4575
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   430
         Width           =   1335
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   3120
         TabIndex        =   26
         Top             =   430
         Width           =   1575
      End
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox CboYear 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrintBordxIBM.frx":0004
         Left            =   3120
         List            =   "frmPrintBordxIBM.frx":0006
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox CboMonth 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrintBordxIBM.frx":0008
         Left            =   3120
         List            =   "frmPrintBordxIBM.frx":0030
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         ItemData        =   "frmPrintBordxIBM.frx":0096
         Left            =   1800
         List            =   "frmPrintBordxIBM.frx":0098
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label lblTo 
         Caption         =   "To Date"
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   430
         Width           =   735
      End
      Begin VB.Label lblFromDate 
         Caption         =   "From Date"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   430
         Width           =   855
      End
      Begin VB.Label lblGrpcompany 
         Caption         =   "Group Company"
         Height          =   255
         Left            =   5880
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblDept 
         Caption         =   "Department"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Month"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Print Bordx - SEB/CIMB/INDAHWATER/AGROBANK/PENGERANG"
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
         TabIndex        =   4
         Top             =   60
         Width           =   11085
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   5175
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9128
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
Attribute VB_Name = "frmPrintBordxIBM"
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
Dim GrpCompany As String
Dim tmpRptFormat As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim tmpRptType As String
Dim tmpRptThphoidType As String
Dim MBMCnt As Integer
Dim CshCnt As Integer
Dim PmntClaimed As Double
Dim PmntApproved As Double
Dim Pmntdiscount As Double

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
Dim LastDate As String
Dim tmpbordxcoveramt As Double


Private m_blnDirection(1 To 13) As Boolean

Private Sub CboMonth_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboMonth, KeyAscii)
End Sub

Private Sub CboYear_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboYear, KeyAscii)
End Sub

Private Sub CmbGrpcomapny_Click()
If CmbGrpcomapny = "IBM" Then
        Label4.Visible = True
        TxtBordxNo.Visible = True
        Label7.Visible = True
        CboMonth.Visible = True
        Label3.Visible = True
        CboYear.Visible = True
        
        txtFrom.Visible = False
        txtTo.Visible = False
        lblTo.Visible = False
        lblFromDate.Visible = False
 
 ElseIf CmbGrpcomapny = "CIMB" Then
        Label4.Visible = False
        TxtBordxNo.Visible = False
        Label7.Visible = False
        CboMonth.Visible = False
        Label3.Visible = False
        CboYear.Visible = False
        
        txtFrom.Visible = True
        txtTo.Visible = True
        lblTo.Visible = True
        lblFromDate.Visible = True
  ElseIf CmbGrpcomapny = "INDAH WATER KONSORTIUM SDN BHD" Then
        Label4.Visible = True
        TxtBordxNo.Visible = True
        Label7.Visible = False
        CboMonth.Visible = False
        Label3.Visible = False
        CboYear.Visible = False
        txtFrom.Visible = False
        txtTo.Visible = False
        lblTo.Visible = False
        lblFromDate.Visible = False
 ElseIf CmbGrpcomapny = "BANK PERTANIAN MALAYSIA BERHAD (AGROBANK)" Then
        Label4.Visible = True
        TxtBordxNo.Visible = True
        Label7.Visible = False
        CboMonth.Visible = False
        Label3.Visible = False
        CboYear.Visible = False
        txtFrom.Visible = False
        txtTo.Visible = False
        lblTo.Visible = False
        lblFromDate.Visible = False
 ElseIf CmbGrpcomapny = "PENGERANG REFINING COMPANY SDN BHD" Then
        Label4.Visible = True
        TxtBordxNo.Visible = True
        Label7.Visible = False
        CboMonth.Visible = False
        Label3.Visible = False
        CboYear.Visible = False
        txtFrom.Visible = False
        txtTo.Visible = False
        lblTo.Visible = False
        lblFromDate.Visible = False
ElseIf CmbGrpcomapny = "SARAWAK ENERGY BERHAD" Or CmbGrpcomapny = "SARAWAK ENERGY BERHAD - RETIREE" Or CmbGrpcomapny = "GREAT EASTERN TAKAFUL BERHAD" Or CmbGrpcomapny = "ORSTED SERVICES MALAYSIA SDN BHD" Or CmbGrpcomapny = "THE SOUTH - EAST ASIAN CENTRAL BANKS (SEACEN) RESEARCH AND TRAINING" Or CmbGrpcomapny = "SCORE ASIA SDN. BHD." Then
        Label4.Visible = True
        TxtBordxNo.Visible = True
        Label7.Visible = False
        CboMonth.Visible = False
        Label3.Visible = False
        CboYear.Visible = False
        txtFrom.Visible = False
        txtTo.Visible = False
        lblTo.Visible = False
        lblFromDate.Visible = False
 End If
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
Private Sub PrintCoverLetter(tmpcovercompany As String, tmpcoveraadd1 As String, tmpcoveraadd2 As String, tmpcoveraadd3 As String, tmpcoveraadd4 As String, tmpcoveraadd5 As String, tmpcoveraadd6 As String, tmpcoverbordxrefno As String, tmpcoverbordxdate As String, tmpbordxtype As String, tmpbordxcoveramt As Double)
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
Dim StrSql2 As String
Dim INVPayToset As New ADODB.Recordset

 Set objword = Nothing
 If objword Is Nothing Then
    Set objword = New Word.Application
    objword.Visible = True
    Set objdoc = objword.Documents.Add(LocCardPath & "Claim\coverletter.dot")
    objdoc.Bookmarks("Add2").Range.Text = Trim(tmpcovercompany) 'MBMGroupCompany
    objdoc.Bookmarks("Add3").Range.Text = Trim(tmpcoveraadd2) 'GRPCoAdd1
    objdoc.Bookmarks("Add4").Range.Text = Trim(tmpcoveraadd3) 'GRPCoAdd2
    objdoc.Bookmarks("Add7").Range.Text = Trim(tmpcoveraadd4) 'GRPCoAdd3
    objdoc.Bookmarks("Add6").Range.Text = Trim(tmpcoveraadd5) 'GRPCoPhone
    objdoc.Bookmarks("Add5").Range.Text = Trim(tmpcoveraadd1) 'GRPCoContact
    objdoc.Bookmarks("Bordxdate").Range.Text = tmpcoverbordxdate 'Bordxdate
    objdoc.Bookmarks("today").Range.Text = Format(Date, "DDmmmmYYYY") 'GETDATE
    StrSql2 = ""
    tmpbordxtype = ""
    StrSql2 = " Select INVPayTo from VIEW_OP_GRPSummary where BordxRefNo = '" & Trim(TxtBordxNo) & "' "
    INVPayToset.Open StrSql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
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

Private Sub StandardPrintCoverLetter()
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
Dim StrSql2 As String
Dim INVPayToset As New ADODB.Recordset
Dim BOR As New ADODB.Recordset
Dim sqlbor As String
sqlbor = ""
tmpbordxtype = ""

 Set objword = Nothing
 If objword Is Nothing Then
    Set objword = New Word.Application
    objword.Visible = True
    Set objdoc = objword.Documents.Add(LocCardPath & "Claim\coverletter.dot")
    
    sqlbor = "Select * from CISDB..VIEW_OP_GRPSummarySEB where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not BOR.EOF
    objdoc.Bookmarks("Add2").Range.Text = IIf(Len(BOR!GroupCompany), Trim(BOR!GroupCompany), "NA")  'MBMGroupCompany
    objdoc.Bookmarks("Add3").Range.Text = IIf(Len(BOR!grpcoadd1), Trim(BOR!grpcoadd1), "NA")  'GRPCoAdd1
    objdoc.Bookmarks("Add4").Range.Text = IIf(Len(BOR!grpcoadd2), Trim(BOR!grpcoadd2), "NA")  'GRPCoAdd2
    objdoc.Bookmarks("Add7").Range.Text = IIf(Len(BOR!grpcoadd3), Trim(BOR!grpcoadd3), "NA") 'GRPCoAdd3
    objdoc.Bookmarks("Add6").Range.Text = IIf(Len(BOR!GRPCoPhone), Trim(BOR!GRPCoPhone), "NA")  'GRPCoPhone
    objdoc.Bookmarks("Add5").Range.Text = IIf(Len(BOR!GRPCoContact), Trim(BOR!GRPCoContact), "NA")  'GRPCoContact
    objdoc.Bookmarks("Bordxdate").Range.Text = BOR!BordxDate 'Bordxdate
    objdoc.Bookmarks("today").Range.Text = Format(Date, "DDmmmmYYYY") 'GETDATE
    tmpbordxtype = tmpbordxtype + BOR!INVPayTo + "|"
    tmpcoverbordxrefno = Trim(BOR!Bordxrefno)
    tmpbordxcoveramt = BOR!CLMBordxAmt
    BOR.MoveNext
    Loop
    BOR.Close
    Set BOR = Nothing
    
    If Trim(tmpbordxtype) = "C|M|" Then
        objdoc.Bookmarks("Bordxrefno").Range.Text = tmpcoverbordxrefno & "" & " (Cashless/Reimbursement) " & "" & "-" & "" & " RM " & "" & IIf(Len(tmpbordxcoveramt), Format(tmpbordxcoveramt, "#,##0.00"), 0)
    ElseIf Trim(tmpbordxtype) = "C|" Then
        objdoc.Bookmarks("Bordxrefno").Range.Text = tmpcoverbordxrefno & "" & " (Cashless) " & "" & "-" & "" & " RM " & "" & IIf(Len(tmpbordxcoveramt), Format(tmpbordxcoveramt, "#,##0.00"), 0)
    ElseIf Trim(tmpbordxtype) = "M|" Then
        objdoc.Bookmarks("Bordxrefno").Range.Text = tmpcoverbordxrefno & "" & " (Reimbursement) " & "" & "-" & "" & " RM " & "" & IIf(Len(tmpbordxcoveramt), Format(tmpbordxcoveramt, "#,##0.00"), 0)
    End If
 End If
End Sub

Public Sub BYSARAWAK(objWorkbook As Excel.Workbook)
Dim StrSql As String
Dim cmd As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim Prm8 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim PVCnt As Integer
Dim tmpTotalCases As String
Dim tmpTotalActual, tmpTotalVisits As String
Dim tmpTotalHos As Double
Dim tmpHOSPayee As String
Dim tmpSubMBMPaid As Double
Dim RecFound As Boolean
Dim tmpSeqNo As String
Dim totalMBMPaid, totalBranchamount As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim rst3 As New ADODB.Recordset
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
Dim Rcount As Integer
Dim tmpMBMGroupCompany, tmpSubsidiary As String
tmpSubsidiary = ""
tmpMBMGroupCompany = ""
tmpcoverbordxdate = ""
                    StrSql = ""
                    StrSql = " Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd4.ActiveConnection = MediConnCISWS
                    cmd4.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd4.CommandType = adCmdStoredProc
                    cmd4.CommandTimeout = 900
                    Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd4.Parameters.Append Prm7
                    Set Prm8 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 9)
                    cmd4.Parameters.Append Prm8
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd4.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalVisits = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    totalBranchamount = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    Do While Not rst2.EOF
                    RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                            If PVCnt = 1 Then
                                objWorksheet.Cells(2, "H") = rst2!MBMGroupCompany
                                tmpMBMGroupCompany = rst2!MBMGroupCompany
                                objWorksheet.Cells(3, "H") = TxtBordxNo
                                objWorksheet.Cells(4, "H") = rst2!BordxDate
                                tmpcoverbordxdate = rst2!BordxDate
                                objWorksheet.Cells(5, "H") = rst2!Subsidiary
                                tmpSubsidiary = rst2!Subsidiary
                                
                            End If
                            objWorksheet.Cells(8 + CDbl(PVCnt), "C") = IIf(Len(rst2!SNo), rst2!SNo, "")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "D") = IIf(Len(rst2!MBMDepartment), rst2!MBMDepartment, "")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "E") = IIf(Len(rst2!MBMCostCenter), rst2!MBMCostCenter, "")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "F") = IIf(Len(rst2!TotalVisits), rst2!TotalVisits, "")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "G") = Format(rst2!invoiceamt, "#,##0.00")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "H") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                            tmpTotalVisits = tmpTotalVisits + IIf(Len(rst2!TotalVisits), rst2!TotalVisits, "")
                            tmpTotalActual = tmpTotalActual + IIf(Len(rst2!invoiceamt), rst2!invoiceamt, "")
                            rst2.MoveNext
                    Loop
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E") = "TOTAL"
                        
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F") = tmpTotalVisits
                        
                        objWorksheet.Cells(10 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "G").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "G") = Format(tmpTotalActual, "#,##0.00")
                        tmpTotalActual = 0
                        PVCnt = 0
                    rst2.Close
                    Set rst2 = Nothing
                
                Set objWorksheet = objWorkbook.Sheets(1)
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from CISDB..VIEW_OP_GRPSummarySEB where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
                    tmpBordxMonth = Month(BOR!BordxDate)
                    tmpBordxMonth = Format(BOR!BordxDate, "mmmm")
                    tmpBordxYear = Year(BOR!BordxDate)
                    objWorksheet.Cells(8, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(10, "K") = BOR!BordxDate
                    objWorksheet.Cells(8, "B") = BOR!subname
                    objWorksheet.Cells(9, "B") = BOR!SubAdd1
                    objWorksheet.Cells(10, "B") = BOR!SubAdd2
                    objWorksheet.Cells(11, "B") = BOR!SubAdd3
                    objWorksheet.Cells(14, "C") = BOR!PIC
                    LastDate = BOR!EOB
                    tmpcoverbordxrefno = BOR!Bordxrefno
                    tmpcoverbordxdate = tmpBordxMonth & " " & tmpBordxYear
                    tmpbordxtype = ""
                    tmpbordxcoveramt = CDbl(BOR!CLMBordxAmt)
                objWorksheet.Cells(25, "I") = BOR!TotalVisits
                objWorksheet.Cells(25, "K") = CDbl(BOR!CLMBordxAmt)
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
                
                If RecFound = True Then
                    StrSql = ""
                    StrSql = " Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 900
                    Set Prm3 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd2.Parameters.Append Prm3
                    Set Prm4 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 10)
                    cmd2.Parameters.Append Prm4
                    rst3.CursorLocation = adUseClient
                    rst3.CursorType = adOpenForwardOnly
                    rst3.CacheSize = 100
                    Set rst3 = cmd2.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    totalBranchamount = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    Do While Not rst3.EOF
                    RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(3)
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                            If PVCnt = 1 Then
                                objWorksheet.Cells(2, "H") = tmpMBMGroupCompany
                                objWorksheet.Cells(3, "H") = TxtBordxNo
                                objWorksheet.Cells(4, "H") = rst3!BordxDate
                                objWorksheet.Cells(5, "H") = tmpSubsidiary
                            End If
                            If rst3!RNo = 1 Then
                                Rcount = Rcount + 1
                                objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Rcount
                                objWorksheet.Cells(8 + CDbl(PVCnt), "B") = IIf(Len(rst3!CasNumber), rst3!CasNumber, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "C") = IIf(Len(rst3!MBMNumber), rst3!MBMNumber, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "D") = IIf(Len(rst3!MBMDepartment), rst3!MBMDepartment, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "E") = IIf(Len(rst3!MBMCostCenter), rst3!MBMCostCenter, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "F") = IIf(Len(rst3!MBMEmployeeNo), rst3!MBMEmployeeNo, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "G") = IIf(Len(rst3!MBMName), rst3!MBMName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "H") = IIf(Len(rst3!MBMIcBcPp), rst3!MBMIcBcPp, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "I") = IIf(Len(rst3!PatientName), rst3!PatientName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "J") = IIf(Len(rst3!DateVisit), rst3!DateVisit, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "K") = IIf(Len(rst3!Diagnosis), rst3!Diagnosis, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "L") = IIf(Len(rst3!CLNType), rst3!CLNType, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "M") = IIf(Len(rst3!CLNName), rst3!CLNName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "N") = IIf(Len(rst3!CLNCity), rst3!CLNCity, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "O") = Format(rst3!CLMBordxAmt, "#,##0.00") 'IIf(Len(rst3!CLMBordxAmt), rst3!CLMBordxAmt, "")
                                totalBranchamount = totalBranchamount + Format(rst3!CLMBordxAmt, "#,##0.00")
                            ElseIf rst3!RNo <> 1 Then
                                Rcount = Rcount + 1
                                objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Rcount
                                objWorksheet.Cells(8 + CDbl(PVCnt), "B") = IIf(Len(rst3!CasNumber), rst3!CasNumber, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "C") = IIf(Len(rst3!MBMNumber), rst3!MBMNumber, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "D") = IIf(Len(rst3!MBMDepartment), rst3!MBMDepartment, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "E") = IIf(Len(rst3!MBMCostCenter), rst3!MBMCostCenter, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "F") = IIf(Len(rst3!MBMEmployeeNo), rst3!MBMEmployeeNo, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "G") = IIf(Len(rst3!MBMName), rst3!MBMName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "H") = IIf(Len(rst3!MBMIcBcPp), rst3!MBMIcBcPp, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "I") = IIf(Len(rst3!PatientName), rst3!PatientName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "J") = IIf(Len(rst3!DateVisit), rst3!DateVisit, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "K") = IIf(Len(rst3!Diagnosis), rst3!Diagnosis, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "L") = IIf(Len(rst3!CLNType), rst3!CLNType, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "M") = IIf(Len(rst3!CLNName), rst3!CLNName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "N") = IIf(Len(rst3!CLNCity), rst3!CLNCity, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "O") = Format(rst3!CLMBordxAmt, "#,##0.00")
                                
                                totalBranchamount = totalBranchamount + Format(rst3!CLMBordxAmt, "#,##0.00")
                             End If
                             If rst3!totalcnt = rst3!RNo Then
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "N").Font.Bold = True
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "N") = "Sub Total"
                                
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "O").Font.Bold = True
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "O") = Format(totalBranchamount, "#,##0.00")
                                totalBranchamount = 0
                                PVCnt = PVCnt + 2
                             End If
                             
                             tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst3!CLMBordxAmt), "#,##0.00")
                             rst3.MoveNext
                    Loop
                        objWorksheet.Cells(10 + CDbl(PVCnt), "N").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "N") = "Total"
                        objWorksheet.Cells(10 + CDbl(PVCnt), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "O").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "O") = Format(tmpTotalActual, "#,##0.00")
                    rst3.Close
                    Set rst3 = Nothing
                RecFound = False
                End If
End Sub

Public Sub BYPENGERANG(objWorkbook As Excel.Workbook)
Dim StrSql As String
Dim cmd As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim Prm8 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim PVCnt As Integer
Dim tmpTotalCases As String
Dim tmpTotalActual, tmpTotalVisits As String
Dim tmpTotalHos As Double
Dim tmpHOSPayee As String
Dim tmpSubMBMPaid As Double
Dim RecFound As Boolean
Dim tmpSeqNo As String
Dim totalMBMPaid, totalBranchamount As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim rst3 As New ADODB.Recordset
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
Dim Rcount As Integer
tmpcoverbordxdate = ""
                    StrSql = ""
                    StrSql = " Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd4.ActiveConnection = MediConnCISWS
                    cmd4.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd4.CommandType = adCmdStoredProc
                    cmd4.CommandTimeout = 900
                    Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd4.Parameters.Append Prm7
                    Set Prm8 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 6)
                    cmd4.Parameters.Append Prm8
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd4.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalVisits = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    totalBranchamount = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    Do While Not rst2.EOF
                    RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                            If PVCnt = 1 Then
                                objWorksheet.Cells(2, "H") = "PENGERANG REFINING COMPANY SDN BHD"
                                objWorksheet.Cells(3, "H") = TxtBordxNo
                                objWorksheet.Cells(4, "H") = rst2!BordxDate
                                tmpcoverbordxdate = rst2!BordxDate
                            End If
                            objWorksheet.Cells(8 + CDbl(PVCnt), "C") = IIf(Len(rst2!SNo), rst2!SNo, "")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "D") = IIf(Len(rst2!MBMGroupCategory), rst2!MBMGroupCategory, "")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "E") = IIf(Len(rst2!MBMCostCenter), rst2!MBMCostCenter, "")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "F") = IIf(Len(rst2!TotalVisits), rst2!TotalVisits, "")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "G") = Format(rst2!invoiceamt, "#,##0.00")
                            objWorksheet.Cells(8 + CDbl(PVCnt), "H") = IIf(Len(rst2!CLNType), rst2!CLNType, "")
                            tmpTotalVisits = tmpTotalVisits + IIf(Len(rst2!TotalVisits), rst2!TotalVisits, "")
                            tmpTotalActual = tmpTotalActual + IIf(Len(rst2!invoiceamt), rst2!invoiceamt, "")
                            rst2.MoveNext
                    Loop
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "E") = "TOTAL"
                        
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "F") = tmpTotalVisits
                        
                        objWorksheet.Cells(10 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "G").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "G") = Format(tmpTotalActual, "#,##0.00")
                        tmpTotalActual = 0
                        PVCnt = 0
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
                    objWorksheet.Cells(10, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(12, "K") = BOR!BordxDate
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
                    StrSql = ""
                    StrSql = " Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 900
                    Set Prm3 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd2.Parameters.Append Prm3
                    Set Prm4 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 8)
                    cmd2.Parameters.Append Prm4
                    rst3.CursorLocation = adUseClient
                    rst3.CursorType = adOpenForwardOnly
                    rst3.CacheSize = 100
                    Set rst3 = cmd2.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    totalBranchamount = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    Do While Not rst3.EOF
                    RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(3)
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                            
                            If PVCnt = 1 Then
                                objWorksheet.Cells(2, "L") = "PENGERANG REFINING COMPANY SDN BHD"
                                objWorksheet.Cells(3, "L") = TxtBordxNo
                                objWorksheet.Cells(4, "L") = rst3!BordxDate
                            End If
                            If rst3!SNo = 1 Then
                                Rcount = Rcount + 1
                                objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Rcount
                                objWorksheet.Cells(8 + CDbl(PVCnt), "B") = IIf(Len(rst3!CasNumber), rst3!CasNumber, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "C") = IIf(Len(rst3!MBMNumber), rst3!MBMNumber, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "D") = IIf(Len(rst3!MBMGroupCategory), rst3!MBMGroupCategory, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "E") = IIf(Len(rst3!MBMCostCenter), rst3!MBMCostCenter, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "F") = IIf(Len(rst3!MBMDivision), rst3!MBMDivision, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "G") = IIf(Len(rst3!MBMEmployeeNo), rst3!MBMEmployeeNo, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "H") = IIf(Len(rst3!MBMName), rst3!MBMName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "I") = IIf(Len(rst3!IcNo), rst3!IcNo, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "J") = IIf(Len(rst3!PatientName), rst3!PatientName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "K") = IIf(Len(rst3!DateVisit), rst3!DateVisit, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "L") = IIf(Len(rst3!Diagnosis), rst3!Diagnosis, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "M") = IIf(Len(rst3!CLNType), rst3!CLNType, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "N") = IIf(Len(rst3!CLNName), rst3!CLNName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "O") = IIf(Len(rst3!CLNState), rst3!CLNState, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "P") = Format(rst3!CLMBordxAmt, "#,##0.00")
                                totalBranchamount = totalBranchamount + Format(rst3!CLMBordxAmt, "#,##0.00")
                            ElseIf rst3!SNo <> 1 Then
                                Rcount = Rcount + 1
                                objWorksheet.Cells(8 + CDbl(PVCnt), "A") = Rcount
                                objWorksheet.Cells(8 + CDbl(PVCnt), "B") = IIf(Len(rst3!CasNumber), rst3!CasNumber, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "C") = IIf(Len(rst3!MBMNumber), rst3!MBMNumber, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "D") = IIf(Len(rst3!MBMGroupCategory), rst3!MBMGroupCategory, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "E") = IIf(Len(rst3!MBMCostCenter), rst3!MBMCostCenter, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "F") = IIf(Len(rst3!MBMDivision), rst3!MBMDivision, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "G") = IIf(Len(rst3!MBMEmployeeNo), rst3!MBMEmployeeNo, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "H") = IIf(Len(rst3!MBMName), rst3!MBMName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "I") = IIf(Len(rst3!IcNo), rst3!IcNo, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "J") = IIf(Len(rst3!PatientName), rst3!PatientName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "K") = IIf(Len(rst3!DateVisit), rst3!DateVisit, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "L") = IIf(Len(rst3!Diagnosis), rst3!Diagnosis, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "M") = IIf(Len(rst3!CLNType), rst3!CLNType, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "N") = IIf(Len(rst3!CLNName), rst3!CLNName, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "O") = IIf(Len(rst3!CLNState), rst3!CLNState, "")
                                objWorksheet.Cells(8 + CDbl(PVCnt), "P") = Format(rst3!CLMBordxAmt, "#,##0.00")
                                totalBranchamount = totalBranchamount + Format(rst3!CLMBordxAmt, "#,##0.00")
                             End If
                             If rst3!totalcnt = rst3!SNo Then
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "O").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "O").Font.Bold = True
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "O") = "Sub Total"
                                
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "P").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "P").Font.Bold = True
                                objWorksheet.Cells(8 + CDbl(PVCnt) + 1, "P") = Format(totalBranchamount, "#,##0.00")
                                totalBranchamount = 0
                                PVCnt = PVCnt + 2
                             End If
                             
                             tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst3!CLMBordxAmt), "#,##0.00")
                             rst3.MoveNext
                    Loop
                        objWorksheet.Cells(10 + CDbl(PVCnt), "O").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "O") = "Total"
                        
                        objWorksheet.Cells(10 + CDbl(PVCnt), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(10 + CDbl(PVCnt), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(10 + CDbl(PVCnt), "P").Font.Bold = True
                        objWorksheet.Cells(10 + CDbl(PVCnt), "P") = Format(tmpTotalActual, "#,##0.00")
                    rst3.Close
                    Set rst3 = Nothing
                RecFound = False
                End If
End Sub
Private Sub DetailPrintAGROBANK_KCN()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
    If objExcel Is Nothing Then
       Set objExcel = CreateObject("Excel.application")
       objExcel.Interactive = False
       objExcel.Visible = True
       objExcel.WindowState = xlMaximized
       Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\AGROBANK_Print_Bordx_KCN_2020.xlt")
    End If
    Call BYAGROBANK(objWorkbook)
    Call PrintCoverLetter("BANK PERTANIAN MALAYSIA BERHAD (AGROBANK)", "Puan Suryati Binti Kamaruddin", "Level 11, Leboh Pasar Besar,", "P.O. Box 10815,", "50726 Kuala Lumpur.", "03-27311600", tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=1, OpenAfterPublish:=True
End Sub
Private Sub DetailPrintSARAWAK()
Const METHOD_NAME = "DetailPrintSARAWAK"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim X As Boolean
Dim Y As Boolean
Dim DocSavePath1 As String
Dim GPDF As String

    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
        objExcel.Interactive = False
        objExcel.Visible = True
        objExcel.WindowState = xlMaximized
        
        If Mid(UCase(Trim(TxtBordxNo)), 1, 3) = "KCN" Then
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\SARAWAK_KCNBillingFormat.xlt")
        Else
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\SARAWAK_BillingFormat.xlt")
        End If
    End If
    Call BYSARAWAK(objWorkbook)
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    
    X = CreatePath(BordxSavePath & "SEB\Invoices\" & LastDate)
    Y = CreatePath(BordxSavePath & "SEB\Invoices\" & LastDate & "\" & "OP")
    DocSavePath1 = BordxSavePath & "SEB\Invoices\" & LastDate & "\OP\"
    GPDF = BordxSavePath & "SEB\Invoices\" & LastDate & "\OP\" & "" & UCase(Trim(TxtBordxNo)) & ".PDF"
    objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=GPDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=50, OpenAfterPublish:=False

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub DetailGE()
Const METHOD_NAME = "DetailGE"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim X As Boolean
Dim Y As Boolean
Dim DocSavePath1 As String
Dim GPDF As String

    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
        objExcel.Interactive = False
        objExcel.Visible = False
        objExcel.WindowState = xlMaximized

        If Mid(UCase(Trim(TxtBordxNo)), 1, 3) = "KCN" Then
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\GEKCN.xlt")
        Else
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\GEKIS.xlt")
        End If
    End If
    Call BYGE(objWorkbook)
    objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=1, OpenAfterPublish:=True
    objWorkbook.Sheets(1).Visible = False
    objExcel.Visible = True
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Call StandardPrintCoverLetter
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub BYGE(objWorkbook As Excel.Workbook)
Dim StrSql As String
Dim cmd As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim Prm8 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim PVCnt As Integer
Dim tmpTotalCases As String
Dim tmpTotalActual, tmpTotalVisits As String
Dim tmpTotalHos As Double
Dim tmpHOSPayee As String
Dim tmpSubMBMPaid As Double
Dim RecFound As Boolean
Dim tmpSeqNo As String
Dim totalMBMPaid, totalBranchamount As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim rst3 As New ADODB.Recordset
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
Dim Rcount As Integer
Dim tmpMBMGroupCompany, tmpSubsidiary As String
tmpSubsidiary = ""
tmpMBMGroupCompany = ""
tmpcoverbordxdate = ""

                    Set objWorksheet = objWorkbook.Sheets(1)
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    sqlbor = ""
                    sqlbor = "Select * from CISDB..VIEW_OP_GRPSummarySEB where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                    Do While Not BOR.EOF
                    objWorksheet.Cells(8, "K") = BOR!Bordxrefno
                    objWorksheet.Cells(10, "K") = BOR!BordxDate
                    objWorksheet.Cells(8, "B") = BOR!GroupCompany
                    objWorksheet.Cells(9, "B") = BOR!grpcoadd1
                    objWorksheet.Cells(10, "B") = BOR!grpcoadd2
                    objWorksheet.Cells(11, "B") = BOR!grpcoadd3
                    objWorksheet.Cells(14, "C") = BOR!GRPCoContact
                    objWorksheet.Cells(25, "I") = BOR!TotalVisits
                    objWorksheet.Cells(25, "K") = CDbl(BOR!CLMBordxAmt)
                    objWorksheet.Cells(25, "B") = "OUTPATIENT MEDICAL COST FOR " + Trim(BOR!GroupCompany) + "-(" + BOR!PAYTo + ")"
                    BOR.MoveNext
                    Loop
                    BOR.Close
                    Set BOR = Nothing
                    
                    
                    StrSql = ""
                    StrSql = " Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 900
                    Set Prm3 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd2.Parameters.Append Prm3
                    Set Prm4 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 12)
                    cmd2.Parameters.Append Prm4
                    rst3.CursorLocation = adUseClient
                    rst3.CursorType = adOpenForwardOnly
                    rst3.CacheSize = 100
                    Set rst3 = cmd2.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    totalBranchamount = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    Do While Not rst3.EOF
                    RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                            If PVCnt = 1 Then
                                objWorksheet.Cells(2, "H") = rst3!MBMGroupCompany
                                objWorksheet.Cells(3, "H") = rst3!refno
                                objWorksheet.Cells(4, "H") = rst3!BordxDate
                            End If
                            Rcount = Rcount + 1
                            objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Rcount
                            objWorksheet.Cells(7 + CDbl(PVCnt), "B") = IIf(Len(rst3!CasNumber), rst3!CasNumber, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "C") = IIf(Len(rst3!CLNType), rst3!CLNType, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "D") = IIf(Len(rst3!MBMName), rst3!MBMName, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "E") = IIf(Len(rst3!MBMNumber), rst3!MBMNumber, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "F") = IIf(Len(rst3!PatientName), rst3!PatientName, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "G") = IIf(Len(rst3!PatientIC), rst3!PatientIC, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "H") = IIf(Len(rst3!Diagnosis), rst3!Diagnosis, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "I") = IIf(Len(rst3!CLNName), rst3!CLNName, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "J") = IIf(Len(rst3!CLNCity), rst3!CLNCity, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "K") = IIf(Len(rst3!DateVisit), rst3!DateVisit, "")
                            objWorksheet.Cells(7 + CDbl(PVCnt), "L") = Format(rst3!CLMBordxAmt, "#,##0.00")
                            
                            totalBranchamount = totalBranchamount + Format(rst3!CLMBordxAmt, "#,##0.00")
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst3!CLMBordxAmt), "#,##0.00")
                             rst3.MoveNext
                    Loop
                        
                        objWorksheet.Cells(8 + CDbl(PVCnt), "K").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "K") = "Sub Total"
                        objWorksheet.Cells(8 + CDbl(PVCnt), "L").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(8 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(8 + CDbl(PVCnt), "L").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "L") = Format(tmpTotalActual, "#,##0.00")
                    
                    
                        PVCnt = PVCnt + 2
                        objWorksheet.Cells(9 + CDbl(PVCnt), "K").Font.Bold = True
                        objWorksheet.Cells(9 + CDbl(PVCnt), "K") = "Total"
                        objWorksheet.Cells(9 + CDbl(PVCnt), "L").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(9 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(9 + CDbl(PVCnt), "L").Font.Bold = True
                        objWorksheet.Cells(9 + CDbl(PVCnt), "L") = Format(tmpTotalActual, "#,##0.00")
                    rst3.Close
                    Set rst3 = Nothing
                RecFound = False
                
End Sub

Public Function CreatePath(NewPath) As Boolean
    'Add a trailing slash if none
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If
End Function
Private Sub DetailPrintPENGERANG()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
        objExcel.Interactive = False
        objExcel.Visible = True
        objExcel.WindowState = xlMaximized
        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\PENGERANG_BillingFormat.xlt")
    End If
    Call BYPENGERANG(objWorkbook)
    Call PrintCoverLetter("PENGERANG REFINING COMPANY SDN BHD", "", "HRM OFFICE, ADMINISTRATION BLOCK B1, FIRST FLOOR,", "PENGERANG INTEGRATED COMPLEX MANAGEMENT OFFICE,", "81600 PENGERANG, JOHOR", "03 23923236", tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=1, OpenAfterPublish:=True
End Sub
Private Sub DetailPrintAGROBANK()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
    If objExcel Is Nothing Then
       Set objExcel = CreateObject("Excel.application")
       objExcel.Interactive = False
       objExcel.Visible = True
       objExcel.WindowState = xlMaximized
       Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\AGROBANK_Print_Bordx_2020.xlt")
    End If
    Call BYAGROBANK(objWorkbook)
    Call PrintCoverLetter("BANK PERTANIAN MALAYSIA BERHAD (AGROBANK)", "Puan Suryati Binti Kamaruddin", "Level 11, Leboh Pasar Besar,", "P.O. Box 10815,", "50726 Kuala Lumpur.", "03-27311600", tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=1, OpenAfterPublish:=True
End Sub
Public Sub BYAGROBANK(objWorkbook As Excel.Workbook)
Dim StrSql As String
Dim cmd As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim cmd5 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim PVCnt As Integer
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
Dim rst3 As New ADODB.Recordset
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
Dim Rcount As Integer
Dim ChkStatus, MTIPstatus As Boolean
Dim Prm7 As New ADODB.Parameter
Dim Prm8 As New ADODB.Parameter
Dim Cnt As Integer

                    'Sheet1
                    tmpbordxcoveramt = "0.00"
                    tmpcoverbordxdate = ""
                    sqlbor = ""
                    sqlbor = " Bordx.Bordxrefno = '" & Trim(TxtBordxNo) & "'"
                    Set cmd4.ActiveConnection = MediConnCISWS
                    cmd4.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd4.CommandType = adCmdStoredProc
                    cmd4.CommandTimeout = 900
                    Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sqlbor)
                    cmd4.Parameters.Append Prm7
                    Set Prm8 = cmd5.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 5)
                    cmd4.Parameters.Append Prm8
                    BOR.CursorLocation = adUseClient
                    BOR.CursorType = adOpenForwardOnly
                    BOR.CacheSize = 100
                    Set BOR = cmd4.Execute
                    
                    ChkStatus = True
                    MTIPstatus = True
                    Cnt = 21
                    Do While Not BOR.EOF
                    Set objWorksheet = objWorkbook.Sheets(1)
                    Set objWorksheet = objWorkbook.Worksheets.Item(1)
                        objWorksheet.Cells(10, "K") = TxtBordxNo
                        objWorksheet.Cells(12, "K") = BOR!BordxDate
                        tmpcoverbordxdate = BOR!BordxDate
                  If Trim(BOR!INVPayTo) = "C" Then
                        If BOR!Claimtype = "MATERNITY-INPATIENT" And MTIPstatus = True Then
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = "INPATIENT MEDICAL COST(CASHLESS)"
                            objWorksheet.Cells(Cnt, "B").Font.Bold = True
                            objWorksheet.Cells(Cnt, "B").Font.Underline = True
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = BOR!ServiceType
                            objWorksheet.Cells(Cnt, "I") = BOR!Cnt
                            objWorksheet.Cells(Cnt, "K") = BOR!CLMBordxAmt
                            MTIPstatus = False
                        ElseIf ChkStatus = True And BOR!Claimtype <> "MATERNITY-INPATIENT" Then
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = "OUTPATIENT MEDICAL COST(CASHLESS)"
                            objWorksheet.Cells(Cnt, "B").Font.Bold = True
                            objWorksheet.Cells(Cnt, "B").Font.Underline = True
                            
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = BOR!ServiceType
                            objWorksheet.Cells(Cnt, "I") = BOR!Cnt
                            objWorksheet.Cells(Cnt, "K") = BOR!CLMBordxAmt
                            ChkStatus = False
                        Else
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = BOR!ServiceType
                            objWorksheet.Cells(Cnt, "I") = BOR!Cnt
                            objWorksheet.Cells(Cnt, "K") = BOR!CLMBordxAmt
                        End If
                    ElseIf Trim(BOR!INVPayTo) = "M" Then
                        If BOR!Claimtype = "MATERNITY-INPATIENT" And MTIPstatus = True Then
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = "INPATIENT MEDICAL COST(REIMBURSEMENT)"
                            objWorksheet.Cells(Cnt, "B").Font.Bold = True
                            objWorksheet.Cells(Cnt, "B").Font.Underline = True
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = BOR!ServiceType
                            objWorksheet.Cells(Cnt, "I") = BOR!Cnt
                            objWorksheet.Cells(Cnt, "K") = BOR!CLMBordxAmt
                            MTIPstatus = False
                        ElseIf ChkStatus = True And BOR!Claimtype <> "MATERNITY-INPATIENT" Then
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = "OUTPATIENT MEDICAL COST(REIMBURSEMENT)"
                            objWorksheet.Cells(Cnt, "B").Font.Bold = True
                            objWorksheet.Cells(Cnt, "B").Font.Underline = True
                            
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = BOR!ServiceType
                            objWorksheet.Cells(Cnt, "I") = BOR!Cnt
                            objWorksheet.Cells(Cnt, "K") = BOR!CLMBordxAmt
                            ChkStatus = False
                        Else
                            Cnt = Cnt + 2
                            objWorksheet.Cells(Cnt, "B") = BOR!ServiceType
                            objWorksheet.Cells(Cnt, "I") = BOR!Cnt
                            objWorksheet.Cells(Cnt, "K") = BOR!CLMBordxAmt
                        End If
                    End If
                    BOR.MoveNext
                    Loop
                    BOR.Close
                    Set BOR = Nothing

                    StrSql = ""
                    StrSql = " Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 900
                    Set Prm3 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd2.Parameters.Append Prm3
                    Set Prm4 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 3)
                    cmd2.Parameters.Append Prm4
                    rst3.CursorLocation = adUseClient
                    rst3.CursorType = adOpenForwardOnly
                    rst3.CacheSize = 100
                    Set rst3 = cmd2.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    Do While Not rst3.EOF
                    RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                            If PVCnt = 1 Then
                                objWorksheet.Cells(2, "I") = "BANK PERTANIAN MALAYSIA BERHAD (AGROBANK)"
                                objWorksheet.Cells(3, "I") = TxtBordxNo
                                objWorksheet.Cells(4, "I") = rst3!BordxDate
                            End If
                            If rst3!RNo = 1 Then
                                Rcount = Rcount + 1
                                objWorksheet.Cells(7 + CDbl(PVCnt), "A") = rst3!RNo
                                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = IIf(Len(rst3!ClaimNo), rst3!ClaimNo, "")
                                objWorksheet.Cells(7 + CDbl(PVCnt), "C") = rst3!Claimtype
                                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = rst3!MemberNo
                                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = rst3!PatientName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = rst3!Relation
                                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = rst3!Sex
                                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = rst3!DOB
                                objWorksheet.Cells(7 + CDbl(PVCnt), "I") = rst3!NRIC
                                objWorksheet.Cells(7 + CDbl(PVCnt), "J") = rst3!EmployeeName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "K") = rst3!EmployeeNo
                                objWorksheet.Cells(7 + CDbl(PVCnt), "L") = rst3!plancode
                                objWorksheet.Cells(7 + CDbl(PVCnt), "M") = rst3!EffectiveDate
                                objWorksheet.Cells(7 + CDbl(PVCnt), "N") = rst3!ClinicCode
                                
                                objWorksheet.Cells(7 + CDbl(PVCnt), "O") = rst3!ClinicName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "P") = rst3!Diagnosis
                                objWorksheet.Cells(7 + CDbl(PVCnt), "Q") = rst3!ServiceType
                                ''objWorksheet.Cells(7 + CDbl(PVCnt), "R") = rst3!TreatmentDetail
                                objWorksheet.Cells(7 + CDbl(PVCnt), "R") = rst3!DateVisit
                                objWorksheet.Cells(7 + CDbl(PVCnt), "S") = rst3!ServiceMonth
                                objWorksheet.Cells(7 + CDbl(PVCnt), "T") = rst3!InvoiceDate
                                objWorksheet.Cells(7 + CDbl(PVCnt), "U") = rst3!CLMBordxAmt
                            ElseIf rst3!RNo <> 1 Then
                                Rcount = Rcount + 1
                                objWorksheet.Cells(7 + CDbl(PVCnt), "A") = rst3!RNo
                                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = IIf(Len(rst3!ClaimNo), rst3!ClaimNo, "")
                                objWorksheet.Cells(7 + CDbl(PVCnt), "C") = rst3!Claimtype
                                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = rst3!MemberNo
                                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = rst3!PatientName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = rst3!Relation
                                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = rst3!Sex
                                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = rst3!DOB
                                objWorksheet.Cells(7 + CDbl(PVCnt), "I") = rst3!NRIC
                                objWorksheet.Cells(7 + CDbl(PVCnt), "J") = rst3!EmployeeName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "K") = rst3!EmployeeNo
                                objWorksheet.Cells(7 + CDbl(PVCnt), "L") = rst3!plancode
                                objWorksheet.Cells(7 + CDbl(PVCnt), "M") = rst3!EffectiveDate
                                objWorksheet.Cells(7 + CDbl(PVCnt), "N") = rst3!ClinicCode
                                objWorksheet.Cells(7 + CDbl(PVCnt), "O") = rst3!ClinicName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "P") = rst3!Diagnosis
                                objWorksheet.Cells(7 + CDbl(PVCnt), "Q") = rst3!ServiceType
                                ''objWorksheet.Cells(7 + CDbl(PVCnt), "R") = rst3!TreatmentDetail
                                objWorksheet.Cells(7 + CDbl(PVCnt), "R") = rst3!DateVisit
                                objWorksheet.Cells(7 + CDbl(PVCnt), "S") = rst3!ServiceMonth
                                objWorksheet.Cells(7 + CDbl(PVCnt), "T") = rst3!InvoiceDate
                                objWorksheet.Cells(7 + CDbl(PVCnt), "U") = rst3!CLMBordxAmt
                            End If
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst3!CLMBordxAmt), "#,##0.00")
                            rst3.MoveNext
                    Loop
                        objWorksheet.Cells(8 + CDbl(PVCnt), "T").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "T").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(8 + CDbl(PVCnt), "T").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(8 + CDbl(PVCnt), "T") = "Total"
                        
                        objWorksheet.Cells(8 + CDbl(PVCnt), "U").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "U").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(8 + CDbl(PVCnt), "U").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(8 + CDbl(PVCnt), "U") = Format(tmpTotalActual, "#,##0.00")
                        
                        
                        PVCnt = PVCnt + 3
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B").Borders(xlEdgeBottom).LineStyle = xlContinuous
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "NOTE:"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "OP : OUTPATIENT"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "IP : INPATIENT"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "DT : DENTAL"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "OC : OPTICAL"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "GP : GENERAL PHYSICIAN"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "SP : SPECIALIST"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "PR : PRE EMPLOYMENT CHECK UP"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "MT :  MATERNITY"
                        PVCnt = PVCnt + 1
                        objWorksheet.Cells(8 + CDbl(PVCnt), "B") = "ME : MEDICAL CHECK UP (ESP)"
                        
                    rst3.Close
                    Set rst3 = Nothing
End Sub
Private Sub DetailPrintINDAHWATER()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
        objExcel.Interactive = False
        objExcel.Visible = True
        objExcel.WindowState = xlMaximized
        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "CISTemplate\INDAHWATER_Print Bordx.xlt")
    End If
    Call BYINDAHWATER(objWorkbook)
    Call PrintCoverLetter("INDAH WATER KONSORTIUM SDN BHD", tmpcoveraadd1, tmpcoveraadd2, tmpcoveraadd3, tmpcoveraadd4, tmpcoveraadd5, tmpcoveraadd6, tmpcoverbordxrefno, tmpcoverbordxdate, tmpbordxtype, tmpbordxcoveramt)
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=1, OpenAfterPublish:=True
End Sub
Public Sub BYINDAHWATER(objWorkbook As Excel.Workbook)
Dim StrSql As String
Dim cmd As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim Prm8 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim TMPWS As Integer
Dim BOR As New ADODB.Recordset
Dim PVCnt As Integer
Dim tmpTotalCases As String
Dim tmpTotalActual As String
Dim tmpTotalHos As Double
Dim tmpHOSPayee As String
Dim tmpSubMBMPaid As Double
Dim RecFound As Boolean
Dim tmpSeqNo As String
Dim totalMBMPaid, totalBranchamount As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim rst3 As New ADODB.Recordset
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
Dim Rcount As Integer
tmpcoverbordxdate = ""
                    StrSql = ""
                    StrSql = " Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd4.ActiveConnection = MediConnCISWS
                    cmd4.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd4.CommandType = adCmdStoredProc
                    cmd4.CommandTimeout = 900
                    Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd4.Parameters.Append Prm7
                    Set Prm8 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 1)
                    cmd4.Parameters.Append Prm8
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
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    totalBranchamount = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    Do While Not rst2.EOF
                    RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                            If PVCnt = 1 Then
                                objWorksheet.Cells(2, "G") = "INDAH WATER KONSORTIUM SDN BHD"
                                objWorksheet.Cells(3, "G") = TxtBordxNo
                                objWorksheet.Cells(4, "G") = rst2!BordxDate
                                tmpcoverbordxdate = rst2!BordxDate
                            End If
                            If rst2!RNo = 1 Then
                                objWorksheet.Cells(11 + CDbl(PVCnt), "C") = IIf(Len(rst2!MBMBranch), rst2!MBMBranch, "")
                                objWorksheet.Cells(11 + CDbl(PVCnt), "D") = rst2!MBMCostCenter
                                objWorksheet.Cells(11 + CDbl(PVCnt), "E") = rst2!MBMDepartment
                                objWorksheet.Cells(11 + CDbl(PVCnt), "F") = Format(rst2!CLMBordxAmt, "#,##0.00")
                                totalBranchamount = totalBranchamount + Format(rst2!CLMBordxAmt, "#,##0.00")
                            ElseIf rst2!RNo <> 1 Then
                                objWorksheet.Cells(11 + CDbl(PVCnt), "D") = rst2!MBMCostCenter
                                objWorksheet.Cells(11 + CDbl(PVCnt), "E") = rst2!MBMDepartment
                                objWorksheet.Cells(11 + CDbl(PVCnt), "F") = Format(rst2!CLMBordxAmt, "#,##0.00")
                                totalBranchamount = totalBranchamount + Format(rst2!CLMBordxAmt, "#,##0.00")
                            End If
                            If rst2!RNo = rst2!Cnt Then
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "C").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "C").Font.Bold = True
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "C") = Trim(rst2!MBMBranch) + CStr(" Total")
                                
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "D").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "D").Font.Bold = True
                                
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "E").Font.Bold = True
                                
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "F").Font.Bold = True
                                objWorksheet.Cells(11 + CDbl(PVCnt) + 1, "F") = Format(totalBranchamount, "#,##0.00")
                                totalBranchamount = 0
                                PVCnt = PVCnt + 2
                            End If
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst2!CLMBordxAmt), "#,##0.00")
                            rst2.MoveNext
                    Loop
                        objWorksheet.Cells(12 + CDbl(PVCnt), "E").Font.Bold = True
                        objWorksheet.Cells(12 + CDbl(PVCnt), "E") = "TOTAL"
                        
                        objWorksheet.Cells(12 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                        objWorksheet.Cells(12 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                        objWorksheet.Cells(12 + CDbl(PVCnt), "F").Font.Bold = True
                        objWorksheet.Cells(12 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                        tmpTotalActual = 0
                        PVCnt = 0
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
                    objWorksheet.Cells(29, "B") = "OUTPATIENT MEDICAL COST FOR DISTRICT" & " " & BOR!MBMGroupCompany
                    tmpcovercompany = IIf(Len(BOR!MBMGroupCompany), BOR!MBMGroupCompany, "")
                    objWorksheet.Cells(12, "B") = Trim(tmpcovercompany)
                    tmpcoveraadd1 = IIf(Len(BOR!GRPCoContact), BOR!GRPCoContact, "NA")
                    objWorksheet.Cells(18, "C") = Trim(tmpcoveraadd1)
                    tmpcoveraadd2 = IIf(Len(BOR!grpcoadd1), BOR!grpcoadd1, "NA")
                    objWorksheet.Cells(13, "B") = Trim(tmpcoveraadd2)
                    tmpcoveraadd3 = IIf(Len(BOR!grpcoadd2), BOR!grpcoadd2, "NA")
                    objWorksheet.Cells(14, "B") = Trim(tmpcoveraadd3)
                    tmpcoveraadd4 = IIf(Len(BOR!grpcoadd3), BOR!grpcoadd3, "NA")
                    objWorksheet.Cells(15, "B") = Trim(tmpcoveraadd4)
                    tmpcoveraadd5 = IIf(Len(BOR!GRPCoPhone), BOR!GRPCoPhone, "NA")
                    objWorksheet.Cells(16, "B") = Trim(tmpcoveraadd5)
                    tmpcoveraadd6 = IIf(Len(BOR!GRPCoFax), BOR!GRPCoFax, "NA")
                    objWorksheet.Cells(17, "B") = Trim(tmpcoveraadd6)
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
                    StrSql = ""
                    StrSql = " Claim.BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "Sp_IndahWaterSummary_2018"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 900
                    Set Prm3 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
                    cmd2.Parameters.Append Prm3
                    Set Prm4 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 2)
                    cmd2.Parameters.Append Prm4
                    rst3.CursorLocation = adUseClient
                    rst3.CursorType = adOpenForwardOnly
                    rst3.CacheSize = 100
                    Set rst3 = cmd2.Execute
                    Diag = ""
                    PVCnt = 0
                    Rcount = 0
                    tmpTotalCases = 0
                    tmpTotalActual = 0
                    tmpTotalHos = 0
                    tmpHOSPayee = ""
                    tmpSubMBMPaid = 0
                    tmpSeqNo = 0
                    totalMBMPaid = 0
                    totalBranchamount = 0
                    PVReimCnt = 0
                    tmpPatientRel = ""
                    Do While Not rst3.EOF
                    RecFound = True
                        PVCnt = PVCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(3)
                        Set objWorksheet = objWorkbook.Worksheets.Item(3)
                            If PVCnt = 1 Then
                                objWorksheet.Cells(2, "I") = "INDAH WATER KONSORTIUM SDN BHD"
                                objWorksheet.Cells(3, "I") = TxtBordxNo
                                objWorksheet.Cells(4, "I") = rst3!BordxDate
                            End If
                            If rst3!RNo = 1 Then
                                Rcount = Rcount + 1
                                objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Rcount
                                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = IIf(Len(rst3!CasNumber), rst3!CasNumber, "")
                                objWorksheet.Cells(7 + CDbl(PVCnt), "C") = rst3!MBMNumber
                                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = rst3!MBMBranch
                                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = rst3!MBMDepartment
                                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = rst3!MBMCostCenter
                                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = rst3!MBMEmployeeNo
                                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = rst3!StaffName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "I") = rst3!MBMIcBcPp
                                objWorksheet.Cells(7 + CDbl(PVCnt), "J") = rst3!PatientName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "K") = rst3!DateVisit
                                objWorksheet.Cells(7 + CDbl(PVCnt), "L") = rst3!Diagnosis
                                objWorksheet.Cells(7 + CDbl(PVCnt), "M") = rst3!CLNType
                                objWorksheet.Cells(7 + CDbl(PVCnt), "N") = rst3!CLNName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "O") = rst3!CLNCity
                                objWorksheet.Cells(7 + CDbl(PVCnt), "P") = Format(rst3!CLMBordxAmt, "#,##0.00")
                                totalBranchamount = totalBranchamount + Format(rst3!CLMBordxAmt, "#,##0.00")
                            ElseIf rst3!RNo <> 1 Then
                                Rcount = Rcount + 1
                                objWorksheet.Cells(7 + CDbl(PVCnt), "A") = Rcount
                                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = IIf(Len(rst3!CasNumber), rst3!CasNumber, "")
                                objWorksheet.Cells(7 + CDbl(PVCnt), "C") = rst3!MBMNumber
                                objWorksheet.Cells(7 + CDbl(PVCnt), "D") = rst3!MBMBranch
                                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = rst3!MBMDepartment
                                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = rst3!MBMCostCenter
                                objWorksheet.Cells(7 + CDbl(PVCnt), "G") = rst3!MBMEmployeeNo
                                objWorksheet.Cells(7 + CDbl(PVCnt), "H") = rst3!StaffName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "I") = rst3!MBMIcBcPp
                                objWorksheet.Cells(7 + CDbl(PVCnt), "J") = rst3!PatientName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "K") = rst3!DateVisit
                                objWorksheet.Cells(7 + CDbl(PVCnt), "L") = rst3!Diagnosis
                                objWorksheet.Cells(7 + CDbl(PVCnt), "M") = rst3!CLNType
                                objWorksheet.Cells(7 + CDbl(PVCnt), "N") = rst3!CLNName
                                objWorksheet.Cells(7 + CDbl(PVCnt), "O") = rst3!CLNCity
                                objWorksheet.Cells(7 + CDbl(PVCnt), "P") = Format(rst3!CLMBordxAmt, "#,##0.00")
                                totalBranchamount = totalBranchamount + Format(rst3!CLMBordxAmt, "#,##0.00")
                            End If
                            If rst3!RNo = rst3!Cnt Then
                                objWorksheet.Cells(7 + CDbl(PVCnt) + 1, "O").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(7 + CDbl(PVCnt) + 1, "O").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(7 + CDbl(PVCnt) + 1, "O").Font.Bold = True
                                objWorksheet.Cells(7 + CDbl(PVCnt) + 1, "O") = "Sub Total"
                                
                                
                                objWorksheet.Cells(7 + CDbl(PVCnt) + 1, "P").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(7 + CDbl(PVCnt) + 1, "P").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(7 + CDbl(PVCnt) + 1, "P").Font.Bold = True
                                objWorksheet.Cells(7 + CDbl(PVCnt) + 1, "P") = Format(totalBranchamount, "#,##0.00")
                                totalBranchamount = 0
                                PVCnt = PVCnt + 2
                            End If
                            tmpTotalActual = CDbl(tmpTotalActual) + Format(CDbl(rst3!CLMBordxAmt), "#,##0.00")
                            rst3.MoveNext
                    Loop
                        objWorksheet.Cells(8 + CDbl(PVCnt), "O").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "O") = "Total"
                        
                        objWorksheet.Cells(8 + CDbl(PVCnt), "P").Font.Bold = True
                        objWorksheet.Cells(8 + CDbl(PVCnt), "P") = Format(tmpTotalActual, "#,##0.00")
                    rst3.Close
                    Set rst3 = Nothing
                RecFound = False
                End If
End Sub

Private Sub DetailPrint()
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
Dim Prm4 As New ADODB.Parameter
Dim CaseCnt As String
Dim copay As String
Dim SumCoPay As Double
Dim IsPharmacy As String

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
                   
               With BOR
                
                If Len(!IsPharmacy) Then
                  IsPharmacy = !IsPharmacy
                  End If
                If IsNull(!IsPharmacy) Then
                   IsPharmacy = "N"
                   End If
                 If !IsPharmacy = "" Then
                   IsPharmacy = "N"
                   End If
                If !IsPharmacy = " " Then
                   IsPharmacy = "N"
                End If
           
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
                        
                        If (Len(!CLMCoPayCLNColl)) Then
                            SumCoPay = SumCoPay + CDbl(!CLMCoPayCLNColl)
                        End If
                        
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "I") = TempCLNName
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "J") = !CLNType
                         objWorksheet.Cells(5 + CDbl(CaseCnt), "K") = IsPharmacy
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "L") = !CLMTotalActual
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "M") = !CLMBordxAmt
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "N") = 0
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "O") = !CLMCoPayCLNColl
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "P") = !CLMBordxAmt
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "Q") = BOR!INVPayTo
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "R") = tmpDateVisit
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "S") = BOR!MCDays
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "T") = !ReasonDesc1
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "U") = !ReasonDesc2
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "V") = !ReasonDesc3
                        objWorksheet.Cells(5 + CDbl(CaseCnt), "W") = !CASDiagnosis5
                        
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
                sqlbor = ""
                
                sqlbor = "Select * from View_CISPrintBordx_Summary   where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
                BOR.Open sqlbor, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
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
                                objWorksheet.Cells(32, "I") = CDbl(CaseCnt)
                                objWorksheet.Cells(32, "K") = CDbl(SumCoPay)
                            'Else
                            '    objWorksheet.Cells(24, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(24, "I"))
                            '    objWorksheet.Cells(24, "L") = CDbl(BOR!CLMBordxAmt)
                            'End If
                        ElseIf Trim(BOR!MBMDepartment) = "FTH" Then
                            'If BOR!CLNType = "G" Then
                                objWorksheet.Cells(26, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(26, "I"))
                                objWorksheet.Cells(26, "K") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(26, "M") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(32, "I") = CDbl(CaseCnt)
                                objWorksheet.Cells(32, "K") = CDbl(SumCoPay)
                                 
                            'Else
                            '    objWorksheet.Cells(26, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(26, "I"))
                            '    objWorksheet.Cells(26, "L") = CDbl(BOR!CLMBordxAmt)
                            'End If
                        ElseIf Trim(BOR!MBMDepartment) = "RETIREE" Then
                            'If BOR!CLNType = "G" Then
                                objWorksheet.Cells(28, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(28, "I"))
                                objWorksheet.Cells(28, "K") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(28, "M") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(32, "I") = CDbl(CaseCnt)
                                objWorksheet.Cells(32, "K") = CDbl(SumCoPay)
                            'Else
                            '    objWorksheet.Cells(28, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(28, "I"))
                            '    objWorksheet.Cells(28, "L") = CDbl(BOR!CLMBordxAmt)
                            'End If
                        Else
                            'If BOR!CLNType = "G" Then
                                objWorksheet.Cells(30, "I") = BOR!TotalCases + CDbl(objWorksheet.Cells(30, "I"))
                                objWorksheet.Cells(30, "K") = CDbl(BOR!CLMBordxAmt)
                                objWorksheet.Cells(30, "M") = CDbl(BOR!CLMBordxAmt)
                                
                                objWorksheet.Cells(32, "I") = CDbl(CaseCnt)
                                objWorksheet.Cells(32, "K") = CDbl(SumCoPay)
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
                'crptSummary.PrintReport
            End If
        End With
        PrintFunction = True
        Exit Function
cancelPrint:
        PrintFunction = False
        CancelPrinting = True
End Function

Private Sub SummaryPrint()
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
Dim prm2 As New ADODB.Parameter

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
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "19")
        cmd.Parameters.Append prm2
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
                                "'" & BOR!grpcoadd1 & "', '" & BOR!grpcoadd2 & "', '" & BOR!grpcoadd3 & "', " & _
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
        If GetFFS(GrpCompany) = True Then
            crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary_FFS.rpt"
        Else
            If tmpRptType = "M" Then
                crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary_Medix.rpt"
            Else
                crptSummary.ReportFileName = crdPath & "\CISReport\ClaimBordx\ClaimBordx_Summary.rpt"
            End If
        End If
    Else
        If GetFFS(GrpCompany) = True Then
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
Const METHOD_NAME = "cmdPrintDetails_Click"
On Error GoTo debugs


    If CmbGrpcomapny = "CIMB" Then
        Call PrintTemplateCIMB
    ElseIf CmbGrpcomapny = "IBM" Then
        If tmpRptFormat = "2" Then
            Call PrintTemplate
        Else
            If CboMonth <> "" And CboYear <> "" Then
                Call DetailPrint
            Else
                MsgBox "Please select Bordx Month and Bordx Year", vbOKOnly + vbCritical, DialogBoxTitle
            End If
        End If
    ElseIf lstPrintBordx.ListItems.Count <> 0 And CmbGrpcomapny = "INDAH WATER KONSORTIUM SDN BHD" Then
        Call DetailPrintINDAHWATER
    ElseIf lstPrintBordx.ListItems.Count <> 0 And Mid(Trim(TxtBordxNo), 1, 3) = "KDC" And CmbGrpcomapny = "BANK PERTANIAN MALAYSIA BERHAD (AGROBANK)" Then
        Call DetailPrintAGROBANK
    ElseIf lstPrintBordx.ListItems.Count <> 0 And Mid(Trim(TxtBordxNo), 1, 3) = "KCN" And CmbGrpcomapny = "BANK PERTANIAN MALAYSIA BERHAD (AGROBANK)" Then
        Call DetailPrintAGROBANK_KCN
    ElseIf lstPrintBordx.ListItems.Count <> 0 And CmbGrpcomapny = "PENGERANG REFINING COMPANY SDN BHD" Then
        Call DetailPrintPENGERANG
    ElseIf lstPrintBordx.ListItems.Count <> 0 And (CmbGrpcomapny = "SARAWAK ENERGY BERHAD" Or CmbGrpcomapny = "SARAWAK ENERGY BERHAD - RETIREE") Then
        Call DetailPrintSARAWAK
    ElseIf lstPrintBordx.ListItems.Count <> 0 And (CmbGrpcomapny = "GREAT EASTERN TAKAFUL BERHAD" Or CmbGrpcomapny = "ORSTED SERVICES MALAYSIA SDN BHD" Or CmbGrpcomapny = "THE SOUTH - EAST ASIAN CENTRAL BANKS (SEACEN) RESEARCH AND TRAINING" Or CmbGrpcomapny = "SCORE ASIA SDN. BHD.") Then
        Call DetailGE
    Else
        MsgBox "Please Select Correct Group Company..!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdPrintSummary_Click()
    If VerifyPrint = True Then
        Call SummaryPrint
    End If
    crptSummary.Reset
End Sub

Private Sub CmdSummaryPreview_Click()
    If VerifyPrint = True Then
        Call SummaryPrint
    End If
    crptSummary.Reset
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPrintBordxIBM")
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
    CmbGrpcomapny.AddItem "INDAH WATER KONSORTIUM SDN BHD"
    CmbGrpcomapny.AddItem "BANK PERTANIAN MALAYSIA BERHAD (AGROBANK)"
    CmbGrpcomapny.AddItem "PENGERANG REFINING COMPANY SDN BHD"
    CmbGrpcomapny.AddItem "SARAWAK ENERGY BERHAD - RETIREE"
    CmbGrpcomapny.AddItem "SARAWAK ENERGY BERHAD"
    CmbGrpcomapny.SelText = "INDAH WATER KONSORTIUM SDN BHD"
    CmbGrpcomapny.AddItem "GREAT EASTERN TAKAFUL BERHAD"
    CmbGrpcomapny.AddItem "ORSTED SERVICES MALAYSIA SDN BHD"
    CmbGrpcomapny.AddItem "THE SOUTH - EAST ASIAN CENTRAL BANKS (SEACEN) RESEARCH AND TRAINING"
    CmbGrpcomapny.AddItem "SCORE ASIA SDN. BHD."
    
    Label7.Visible = False
    CboMonth.Visible = False
    Label3.Visible = False
    CboYear.Visible = False
    txtFrom.Visible = False
    txtTo.Visible = False
    lblTo.Visible = False
    lblFromDate.Visible = False
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
lstPrintBordx.ListItems.Clear

If Trim(TxtBordxNo) <> "" And CmbGrpcomapny = "CIMB" Then
    If Trim(txtTo) = "" Or Trim(txtFrom) = "" Then
        MsgBox "Please enter From Date and To Date", vbOKOnly + vbCritical, DialogBoxTitle
         Screen.MousePointer = Default
         Exit Sub
    Else
        strAppend = " And MBMGroupCompany = 'CIMB BANK BERHAD'and BordxDate >='" & Format(txtFrom.Text, "YYYY/MM/DD") & "' and BordxDate <='" & Format(txtTo.Text, "YYYY/MM/DD") & "' "
    End If
ElseIf Trim(TxtBordxNo) <> "" And CmbGrpcomapny = "IBM" Then
        If Trim(TxtBordxNo) = "" Then
        MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
        TxtBordxNo.SetFocus
         Screen.MousePointer = Default
         Exit Sub
    Else
        ScanCode = TxtBordxNo
        If Mid(ScanCode, 1, 2) = "LO" Then
            TxtBordxNo = ""
        Else
            strAppend = " And BordxRefNo = '" & ChkStr(Trim(TxtBordxNo)) & "'"
        End If
    End If
ElseIf Trim(TxtBordxNo) <> "" And CmbGrpcomapny <> "IBM" Then
''= "INDAH WATER KONSORTIUM SDN BHD" Or CmbGrpcomapny = "BANK PERTANIAN MALAYSIA BERHAD (AGROBANK)" Or CmbGrpcomapny = "PENGERANG REFINING COMPANY SDN BHD" Or CmbGrpcomapny = "SARAWAK ENERGY BERHAD - RETIREE" Or CmbGrpcomapny = "SARAWAK ENERGY BERHAD"
    strAppend = " And BordxRefNo = '" & ChkStr(Trim(TxtBordxNo)) & "'"
Else
    MsgBox "Please enter Bordx Reference No..!", vbOKOnly + vbCritical, DialogBoxTitle
    TxtBordxNo.SetFocus
    Exit Sub
End If
CmdClear.Enabled = False
CmdPrintSummary.Enabled = True
cmdPrintDetails.Enabled = True
CmdSummaryPreview.Enabled = True
cmdDetailPreview.Enabled = True
If Trim(strAppend) <> "" Then
Call PrintBordxListing(strAppend)
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
                            GrpCompany = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
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
Dim StrSql As String
    StrSql = " update Bordx set " & _
             " BordxPrintedDate = '" & Format(Date, "YYYY/MM/DD") & "', " & _
             " BORDXPrintMonth = '" & ChkStr(Trim(CboMonth)) & "', " & _
             " BORDXPrintYear = '" & Trim(CboYear) & "', " & _
             " BordxPrintedStatus = '1', " & _
             " BordxPrintedBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(TxtBordxNo) & "'"
    MediConnCISWS.Execute StrSql
End Sub

Private Sub GetBordxInfo()
Dim strsqlBordx As String
Dim BordxREc As New ADODB.Recordset
Dim tmpGrpCom As String
    
    tmpGrpCom = ""
    tmpRptByDept = 0
    tmpPrintStatus = 0
    CboMonth.Enabled = False
    CboYear.Enabled = False
    strsqlBordx = "SELECT GroupCompany, RptTypeByDept,  " & _
                  "BordxPrintedStatus, BORDXPrintMonth, BORDXPrintYear " & _
                  " From Bordx Where BordxRefNo='" & Trim(TxtBordxNo) & "'"
    BordxREc.Open strsqlBordx, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
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

Private Function GetFFS(GrpCompany As String) As Boolean
Dim GrpRec As New ADODB.Recordset
Dim GrpSql As String
Dim tmpGrpMcoStatus As String
    GetFFS = False
    tmpGrpMcoStatus = "F"
    If Trim(GrpCompany) <> "" Then
        GrpSql = "Select GRPMCOStatus from GRPCompany where GRPCoName='" & GrpCompany & "'"
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

Private Sub PrintTemplateCIMB()

Dim PVCnt As Integer
Dim tmpSeqNo As Integer
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim StrSql As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim GrpRec As New ADODB.Recordset
Dim GrpSql As String
Dim CashBordx As String

Dim TRecords As Integer
Dim TMembers As Integer
Dim Tapproved As Double
Dim TRecover As Double



PmntClaimed = 0
PmntApproved = 0
Pmntdiscount = 0

TRecords = 0
Tapproved = 0
TRecover = 0
TMembers = 0

CashBordx = ""
Screen.MousePointer = vbHourglass
    
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
                objExcel.WindowState = xlMaximized
                'Set objWorkbook = objExcel.Workbooks.Add("C:\Documents and Settings\voon\Desktop\PayadviceCIMB.xlt")
                'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PayadviceCIMB.xltx")
                Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PayadviceCIMB.xlt")
                Set objWorksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
            MBMCnt = 0
          
            tmpSeqNo = 0
            StrSql = ""
            'strsql = " and  MBMGroupCompany='CIMB BANK BERHAD' and BordxDate >= '2013-11-01'and BordxDate <= '2013-11-30'"
            StrSql = " And MBMGroupCompany = 'CIMB BANK BERHAD'and BordxDate >='" & Format(txtFrom.Text, "YYYY/MM/DD") & "' and BordxDate <='" & Format(txtTo.Text, "YYYY/MM/DD") & "' "
            Set cmd.ActiveConnection = MediConnCISWS
            cmd.CommandText = "ClaimBordxModule_CIMB"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrSql)
            cmd.Parameters.Append Prm1
            Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 16)
            cmd.Parameters.Append prm2
            rst2.CursorLocation = adUseClient
            rst2.CursorType = adOpenForwardOnly
            rst2.CacheSize = 100
            Set rst2 = cmd.Execute
            PVCnt = 0
           
            tmpSeqNo = 0
             Do While Not rst2.EOF
                  PVCnt = PVCnt + 1
                  Set objWorksheet = objWorkbook.Sheets(1)
                  objExcel.DisplayAlerts = False
                  Set objWorksheet = objWorkbook.Worksheets.Item(1)
                  tmpSeqNo = CDbl(tmpSeqNo) + 1
                  objWorksheet.Cells(7 + CDbl(PVCnt), "A") = PVCnt
                  
                             If Trim(rst2!INVPayTo) = Trim("M") Then
                                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = "CIMBMEDIX-Reimbursement.txt"
                                objWorksheet.Cells(7 + CDbl(PVCnt), "E") = "N/A"
                                Call PrintReimbrusmentDetails(rst2!Bordxrefno, objWorkbook)
                             Else
                                objWorksheet.Cells(7 + CDbl(PVCnt), "B") = "CIMBMEDIX-Cashless.txt"
                                GrpSql = ""
                                GrpSql = "Select Count (distinct(CISWorksheet.dbo.Claim.ClaimOpenBy)) as TotalProviders from CISWorksheet.dbo.Claim where  CISWorksheet.dbo.Claim.BordxRefNo='" & rst2!Bordxrefno & "'"
                                GrpRec.Open GrpSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                                objWorksheet.Cells(7 + CDbl(PVCnt), "F") = IIf(Len(GrpRec!TotalProviders), GrpRec!TotalProviders, "N/A")
                                GrpRec.Close
                                'add the cashless claims to some value and last to call the method.
                                CashBordx = CashBordx + "," + rst2!Bordxrefno
                             End If
                             objWorksheet.Cells(7 + CDbl(PVCnt), "D") = rst2!Bordxrefno
                             objWorksheet.Cells(7 + CDbl(PVCnt), "E") = rst2!BordxNoClaims
                             objWorksheet.Cells(7 + CDbl(PVCnt), "C") = rst2!BordxDate
                             GrpSql = ""
                             GrpSql = "Select Count (distinct(HISDB.dbo.MBMOthers.MBMNumber)) as TotalMembers from HISDB.dbo.MBMOthers inner join " & _
                                      "CISDB.dbo.CISCas on CISDB.dbo.CISCas.MBMNumber=HISDB.dbo.MBMOthers.MBMNumber inner join CISWorksheet.dbo.Claim on CISWorksheet.dbo.Claim.CASNumber = CISDB.DBO.CISCas.CASNumber " & _
                                      "where CISWorksheet.dbo.Claim.BordxRefNo='" & rst2!Bordxrefno & "'"
                             GrpRec.Open GrpSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                             objWorksheet.Cells(7 + CDbl(PVCnt), "G") = IIf(Len(GrpRec!TotalMembers), GrpRec!TotalMembers, "N/A")
                             TMembers = TMembers + CInt(GrpRec!TotalMembers)
                             GrpRec.Close
                             objWorksheet.Cells(7 + CDbl(PVCnt), "H") = rst2!BordxClaimAmt
                             objWorksheet.Cells(7 + CDbl(PVCnt), "J") = rst2!BordxClaimAmt
                             
                             TRecords = TRecords + CInt(rst2!BordxNoClaims)
                             Tapproved = Tapproved + CDbl(rst2!BordxClaimAmt)
                             
                        rst2.MoveNext
                    Loop
                    
                     objWorksheet.Cells(9 + CDbl(PVCnt), "D") = "Total"
                     objWorksheet.Cells(9 + CDbl(PVCnt), "E") = TRecords
                     objWorksheet.Cells(9 + CDbl(PVCnt), "G") = TMembers
                     objWorksheet.Cells(9 + CDbl(PVCnt), "H") = Tapproved
                     objWorksheet.Cells(9 + CDbl(PVCnt), "J") = Tapproved
                      
                    rst2.Close
                    Set rst2 = Nothing
                    
                    'Call method for Cashless by Provider
                    Dim arrBordx() As String
                    arrBordx() = Split(CashBordx, ",")
                    
                    Dim arrCnt As Integer
                    Dim i As Integer
                    arrCnt = UBound(arrBordx)
                    CshCnt = 0
                    For i = 0 To arrCnt Step 1
                    If Len(arrBordx(i)) > 1 Then
                    Call PrintCashLessDetails(arrBordx(i), objWorkbook)
                    End If
                    Next
                    
                     Set objWorksheet = objWorkbook.Worksheets.Item(2)
                     objWorksheet.Cells(13 + CDbl(MBMCnt), "D") = "Total"
                     objWorksheet.Cells(13 + CDbl(MBMCnt), "E") = PmntClaimed
                     objWorksheet.Cells(13 + CDbl(MBMCnt), "F") = PmntApproved
                     objWorksheet.Cells(13 + CDbl(MBMCnt), "G") = Pmntdiscount
                    
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
        
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        Exit Sub
End Sub
Public Sub PrintReimbrusmentDetails(BordxNo As String, objWorkbook As Excel.Workbook)
Dim objWorksheet As Excel.Worksheet
Dim MemberWorksheet As Excel.Worksheet
Dim rst2 As New ADODB.Recordset
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim StrSql As String
Dim SubTotal As Double
Dim TotalApp As Double
Dim TotalDisc As Double
Dim MbmActual As Double
Dim MbmApproved As Double

                    StrSql = ""
                    StrSql = " CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(BordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 300
                    Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, StrSql)
                    cmd2.Parameters.Append Prm3
                    Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 3000, 1)
                    cmd2.Parameters.Append Prm4
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd2.Execute
                    
                    SubTotal = 0
                    TotalApp = 0
                    TotalDisc = 0
                    MbmActual = 0
                    MbmApproved = 0
                    
                    Do While Not rst2.EOF
                        MBMCnt = MBMCnt + 1
                        Set objWorksheet = objWorkbook.Sheets(2)
                        Set objWorksheet = objWorkbook.Worksheets.Item(2)
                        objWorksheet.Cells(9 + CDbl(MBMCnt), "A") = MBMCnt
                          objWorksheet.Cells(9 + CDbl(MBMCnt), "b") = rst2!MBMName
                          If Trim(rst2!CLNType) = Trim("G") Then
                            objWorksheet.Cells(9 + CDbl(MBMCnt), "C") = "General Practitioner"
                          Else
                            objWorksheet.Cells(9 + CDbl(MBMCnt), "C") = "Specialist "
                          End If
                          objWorksheet.Cells(9 + CDbl(MBMCnt), "D") = CStr(rst2!DateVisit)
                          objWorksheet.Cells(9 + CDbl(MBMCnt), "E") = rst2!CLMTotalActual
                          objWorksheet.Cells(9 + CDbl(MBMCnt), "F") = rst2!CLMTotalActual
                          objWorksheet.Cells(9 + CDbl(MBMCnt), "G") = rst2!CLMCoPayColl
                         
                         Set MemberWorksheet = objWorkbook.Sheets(4)
                         Set MemberWorksheet = objWorkbook.Worksheets.Item(4)
                         
                         MemberWorksheet.Cells(7 + CDbl(MBMCnt), "A") = MBMCnt
                         MemberWorksheet.Cells(7 + CDbl(MBMCnt), "B") = rst2!MBMNumber
                         MemberWorksheet.Cells(7 + CDbl(MBMCnt), "C") = rst2!MBMEmployeeNo
                         MemberWorksheet.Cells(7 + CDbl(MBMCnt), "D") = rst2!MBMName
                         MemberWorksheet.Cells(7 + CDbl(MBMCnt), "E") = rst2!CasNumber
                         MemberWorksheet.Cells(7 + CDbl(MBMCnt), "F") = rst2!CLMTotalActual
                         MemberWorksheet.Cells(7 + CDbl(MBMCnt), "G") = rst2!CLMTotalActual
                                
                         MbmActual = CDbl(MbmActual) + CDbl(rst2!CLMTotalActual)
                         
                         PmntClaimed = CDbl(PmntClaimed) + CDbl(rst2!CLMTotalActual)
                         PmntApproved = CDbl(PmntApproved) + CDbl(rst2!CLMTotalActual)
                         Pmntdiscount = CDbl(Pmntdiscount) + CDbl(rst2!CLMCoPayColl)
                         
                        rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
End Sub

Public Sub PrintCashLessDetails(BordxNo As String, objWorkbook As Excel.Workbook)
Dim objWorksheet As Excel.Worksheet
Dim MemberWorksheet As Excel.Worksheet
Dim rst2 As New ADODB.Recordset
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim StrSql As String
Dim SubTotal As Double
Dim TotalApp As Double
Dim TotalDisc As Double
Dim MbmActual As Double
Dim MbmApproved As Double

                    StrSql = ""
                    StrSql = " CISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(BordxNo) & "'"
                    Set cmd2.ActiveConnection = MediConnCISWS
                    cmd2.CommandText = "CLM_PayAdviceCLNBSN"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 300
                    Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, StrSql)
                    cmd2.Parameters.Append Prm3
                    Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 3000, 1)
                    cmd2.Parameters.Append Prm4
                    rst2.CursorLocation = adUseClient
                    rst2.CursorType = adOpenForwardOnly
                    rst2.CacheSize = 100
                    Set rst2 = cmd2.Execute
                    
                    SubTotal = 0
                    TotalApp = 0
                    TotalDisc = 0
                    MbmActual = 0
                    MbmApproved = 0
                    
                    Set objWorksheet = objWorkbook.Sheets(2)
                    Set objWorksheet = objWorkbook.Worksheets.Item(2)
                    objWorksheet.Cells(11 + CDbl(MBMCnt), "A") = "Provider Claims"
                    
                    Do While Not rst2.EOF
                        MBMCnt = MBMCnt + 1
                        CshCnt = CshCnt + 1
                        
                        
                        objWorksheet.Cells(11 + CDbl(MBMCnt), "A") = MBMCnt
                          objWorksheet.Cells(11 + CDbl(MBMCnt), "b") = rst2!MBMName
                          If Trim(rst2!CLNType) = Trim("G") Then
                            objWorksheet.Cells(11 + CDbl(MBMCnt), "C") = "General Practitioner"
                          Else
                            objWorksheet.Cells(11 + CDbl(MBMCnt), "C") = "Specialist "
                          End If
                          objWorksheet.Cells(11 + CDbl(MBMCnt), "D") = CStr(rst2!DateVisit)
                          objWorksheet.Cells(11 + CDbl(MBMCnt), "E") = rst2!CLMTotalActual
                          objWorksheet.Cells(11 + CDbl(MBMCnt), "F") = rst2!CLMTotalActual
                          objWorksheet.Cells(11 + CDbl(MBMCnt), "G") = rst2!CLMCoPayColl
                          
                         PmntClaimed = CDbl(PmntClaimed) + CDbl(rst2!CLMTotalActual)
                         PmntApproved = CDbl(PmntApproved) + CDbl(rst2!CLMTotalActual)
                         Pmntdiscount = CDbl(Pmntdiscount) + CDbl(rst2!CLMCoPayColl)
                         
                         'worksheet Provider code
                         Set MemberWorksheet = objWorkbook.Sheets(3)
                         Set MemberWorksheet = objWorkbook.Worksheets.Item(3)
                         
                         MemberWorksheet.Cells(9 + CDbl(CshCnt), "A") = CshCnt
                         MemberWorksheet.Cells(9 + CDbl(CshCnt), "B") = rst2!CLNCode
                         MemberWorksheet.Cells(9 + CDbl(CshCnt), "C") = rst2!CLNName
                         MemberWorksheet.Cells(9 + CDbl(CshCnt), "D") = rst2!CasNumber
                         MemberWorksheet.Cells(9 + CDbl(CshCnt), "E") = rst2!CLMTotalActual
                         MemberWorksheet.Cells(9 + CDbl(CshCnt), "F") = rst2!CLMTotalActual
                                
                         MbmActual = CDbl(MbmActual) + CDbl(rst2!CLMTotalActual)
                        rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
    
End Sub

Private Sub PrintTemplate()
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
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub



Private Sub PrintTemplateSummary()
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
        'If objExcel Is Nothing Then
        '    Set objExcel = CreateObject("Excel.application")
        '        objExcel.Interactive = False
        '    If objExcel.Visible = False Then
        '        objExcel.Visible = True
        '    End If
        '    objExcel.WindowState = xlMaximized
        '    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Airod OP Template.xls")
        '    'Set objWorkbook = objExcel.Workbooks.Add("C:\Template\Airod OP Template.xls")
            Set objWorksheet = objWorkbook.Sheets(2)
        '    objExcel.DisplayAlerts = False
        '    Set objWorksheet = objWorkbook.Worksheets.Item(1)
        'End If

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
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub

