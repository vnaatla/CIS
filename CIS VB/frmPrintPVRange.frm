VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintPVRange 
   Caption         =   "Print PV Range"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   555
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7935
      Begin VB.OptionButton optNS 
         Caption         =   "CB - NS Citi2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton optNSMAA 
         Caption         =   "CC - NS MAA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optSihat 
         Caption         =   "CA - Sihat Citi1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboAccount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtPVNoTo 
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox txtPVNoFr 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "PV No  From"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   7935
      Begin VB.CommandButton cmdprintPDF 
         Caption         =   "Print PDF"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   22
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   21
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6840
         TabIndex        =   2
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtPrintingStatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "Printing Not In Progress !"
         Top             =   195
         Width           =   2175
      End
      Begin VB.Label lblTotalChqList 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   195
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotalchqSel 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   195
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPymtType 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   195
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Printing Status"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lsPVMBMListing 
      Height          =   6255
      Left            =   0
      TabIndex        =   20
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PV Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Cheque Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "FinalApp"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Case Number"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsPVListing 
      Height          =   6255
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PV Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Cheque Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ClinicName"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ClinicCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "CLNAdd1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "CLNAdd2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "CLNAdd3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "FinalApp"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "MCOProcFee"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Print Cheque Range"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmPrintPVRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public chqMode As String
Dim clickStatus As String
Dim strAppend As String

Dim tmpTotalBillAmt As Double
Dim tmpTotalFFSAmt As Double
Dim tmp_Rec As Double
Dim tmpTotalAmt As Double
Dim tmpTotalCopayment As Double
Dim tmpTotalClaimAmt As Double
Dim tmpTotalClaimAmtSub As Double
Dim tmpNewRow As Double
Dim PrintPDFstatus As Boolean


Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

Const METHOD_NAME = "cmdPrint_Click"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
    Screen.MousePointer = vbHourglass
    PrintPDFstatus = False
    If (lsPVMBMListing.ListItems.Count > 0 And chqMode = "MedixPayMBM") Or (lsPVListing.ListItems.Count > 0 And chqMode = "MedixPayCLN") Then
        If MsgBox("Do you want to print now?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
            RecFound = False
            txtPrintingStatus = "Printing In Progress!"
            CmdPrint.Enabled = False
            CmdExit.Enabled = False
            
                If chqMode = "MedixPayMBM" Then
                    Call PVPrintingMBM
                ElseIf chqMode = "MedixPayCLN" Then
                    Call PVPrintingCLN_New
                End If
        End If
    Else
        MsgBox "Please select PV to print", vbOKOnly + vbCritical
    End If
    Screen.MousePointer = vbDefault
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub PVPrintingMBM()

Const METHOD_NAME = "PVPrintingMBM"
On Error GoTo debugs


'On Error Resume Next
Dim X As Boolean
Dim DocSavePath As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim RecordFound As Boolean
Dim PVRec As New ADODB.Recordset
Dim RecFound As Boolean
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Cnt As Integer
Dim MBMSql As String
Dim MBM As New ADODB.Recordset
Dim PatientName As String
Dim tmpMBMName As String
Dim PatientSql As String
Dim PatientDB As New ADODB.Recordset
Dim ReceiptNo As String
Dim RecSql As String
Dim REC As New ADODB.Recordset
Dim totAmtPaid As Double
Dim totPostage As Double
Dim totPayeble2M  As Double
Dim tmpBankName As String
Dim tmpBankACCNo As String
Dim tmpPVNo As String
Dim tmppvnorpdf As String
Dim TmpPVCode As String
    TmpPVCode = ""
    strAppend = ""
    tmpPVNo = TxtPVNoFr
 
    Do While Not tmpPVNo > TxtPVNoTo
        StrSql = ""
        StrSql = "SELECT CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISDB.dbo.CISCas.DateVisit, CISDB.dbo.CISCas.Remarks, " & _
                "CISDB.dbo.CISCas.Pending, CISDB.dbo.CISCas.Claimability, CISDB.dbo.CISCas.CASStatus, dbo.Payment2MBM.PVNo, dbo.Payment2MBM.PVMBMIndex, " & _
                "dbo.Payment2MBM.PVDate, dbo.Payment2MBM.PVChequeNo, dbo.Payment2MBM.PVChequeDate, dbo.Payment2MBM.PAYAccPeriodMth, " & _
                "dbo.Payment2MBM.MBMPaidAmt, dbo.Payment2MBM.PAYAccPeriodYr, dbo.Payment2MBM.PVPayMBMStatus, dbo.Payment2MBM.Pay2MBMLastUpdateDate, " & _
                "dbo.Payment2MBM.Pay2MBMLastUpdateUser, HISDB.dbo.CLN.CLNName, CISDB.dbo.CISCas.ClinicCode, CISWorksheet.dbo.Claim.DebitCreditRefNo, " & _
                "CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVEnteredBy, CISWorksheet.dbo.ClaimInvoice.BillNo, " & _
                "CISWorksheet.dbo.Claim.CASNumber, case when coalesce(CISWorksheet.dbo.claim.BordxRefNo,'')='' then CISWorksheet.dbo.claim.BTBGBordxRefNo else CISWorksheet.dbo.claim.BordxRefNo end as BordxRefNo, CISWorksheet.dbo.Claim.CLMPayableByMedix2M, " & _
                "CISWorksheet.dbo.Claim.ClaimCloseDate, dbo.Payment2MBM.PVPayto, CISWorksheet.dbo.Bordx.GroupCompany, dbo.Payment2MBM.PVPayeeBank, " & _
                "dbo.Payment2MBM.PVPayeeACNo, dbo.Payment2MBM.Remark, dbo.Payment2MBM.PVCLNChequeAmt, CISDB.dbo.CISCas.OPBatchNo, " & _
                "CISWorksheet.dbo.Claim.CLMPostageFee, CISWorksheet.dbo.Claim.CLMNonGPFee, case When MBMCName IS NULL Then MBMCrossReference.MBMName else MBMCName END AS MBMName, " & _
                              "CISWorksheet.dbo.Claim.ClaimBordxStatus, dbo.Payment2MBM.PVWordingOne, dbo.Payment2MBM.PVWordingTwo, CISWorksheet.dbo.Claim.CLMFinalApprove, " & _
                              "HISDB.dbo.MBMOthersTwo.MBMBankACNo, HISDB.dbo.MBMOthersTwo.MBMBankName, CISWorksheet.dbo.Claim.CLMTotalActual, " & _
                              "HISDB.dbo.MBMOthers.MBMGroupCompany, CISDB.dbo.CISCas.InvoiceNo, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.MBMOthers.MBMCostCenter, " & _
                              "HISDB.dbo.MBMOthers.MBMSubDivision , HISDB.dbo.MBMOthers.MBMDivision, HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMEmployeeNo " & _
                "FROM dbo.Payment2MBM RIGHT OUTER JOIN CISDB.dbo.CISCas LEFT OUTER JOIN HISDB.dbo.MBMOthersTwo RIGHT OUTER JOIN " & _
                              "HISDB.dbo.MBMOthers INNER JOIN " & _
                              "HISMaintenance.dbo.MBMCrossReference ON HISDB.dbo.MBMOthers.MBMNumber = HISMaintenance.dbo.MBMCrossReference.MBMNumber ON " & _
                              "HISDB.dbo.MBMOthersTwo.MBMNumber = HISMaintenance.dbo.MBMCrossReference.MBMNumber ON " & _
                              "HISMaintenance.dbo.MBMCrossReference.MBMNumber = CISDB.dbo.CISCas.MBMnumber left outer join HISDB..MBMCoveredPersons ON MBMCoveredPersons.MBMCNumber=ciscas.MBMnumber and MBMCoveredPersons.MBMCCoverID=ciscas.CoverID RIGHT OUTER JOIN " & _
                              "CISWorksheet.dbo.Claim LEFT OUTER JOIN CISWorksheet.dbo.Bordx ON CISWorksheet.dbo.Bordx.BordxRefNo = case when coalesce(claim.BordxRefNo,'')='' then claim.BTBGBordxRefNo else claim.BordxRefNo end LEFT OUTER JOIN " & _
                              "CISWorksheet.dbo.ClaimInvoice ON CISWorksheet.dbo.Claim.CASNumber = CISWorksheet.dbo.ClaimInvoice.CASNumber ON " & _
                              "CISWorksheet.dbo.Claim.CASNumber = CISDB.dbo.CISCas.CASNumber ON CISWorksheet.dbo.Claim.CASNumber = dbo.Payment2MBM.CASNumber RIGHT OUTER JOIN " & _
                              "HISDB.dbo.CLN ON CISDB.dbo.CISCas.ClinicCode = HISDB.dbo.CLN.CLNCode " & _
                              "WHERE     (CISWorksheet.dbo.ClaimInvoice.INVPayTo = 'M') AND (CISDB.dbo.CISCas.CASStatus <> 'D') AND " & _
                              "(CISWorksheet.dbo.Claim.CLMPayableByMedix2M + CISWorksheet.dbo.Claim.CLMNonGPFee <> 0) And dbo.Payment2MBM.PVNo = '" & Trim(tmpPVNo) & "'"
               
        PVRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        
        i = 0
        Do While Not PVRec.EOF
            If objExcel Is Nothing Then
                Set objExcel = CreateObject("Excel.application")
                    objExcel.Interactive = False
                    objExcel.Visible = False
                objExcel.WindowState = xlMaximized
                Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-ClinicMBM.xlt")
                Set objWorksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objWorksheet = objWorkbook.Worksheets.Item(1)
            End If
                
            On Error GoTo HANDLE_ERROR
            If i = 0 Then
                Cnt = 1
                MBMSql = "SELECT MBMName, MBMAddress1,MBMAddress2,MBMAddress3, CTYCode, MBMPostCode, STACode" & _
                    " FROM MBM WHERE MBMNumber='" & PVRec!MBMNumber & "'"
                MBM.Open MBMSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                PatientName = IIf(Len(MBM!MBMName), MBM!MBMName, "")
                tmpMBMName = IIf(Len(MBM!MBMName), MBM!MBMName, "")
                'If PVRec!MBMEmployeeNo <> "" Then
                '    objWorksheet.Cells(8, "C") = PVRec!PVPayto & " " & "(" & PVRec!MBMEmployeeNo & ")"
                'Else
                    objWorksheet.Cells(8, "C") = PVRec!PVPayto
                'End If
                If Len(PVRec!MBMGroupCategory) Then
                    objWorksheet.Cells(9, "C") = PVRec!GroupCompany & " (" & PVRec!MBMGroupCategory & ")"
                Else
                    objWorksheet.Cells(9, "C") = PVRec!GroupCompany
                End If
                objWorksheet.Cells(10, "C") = MBM!MBMAddress1
                objWorksheet.Cells(11, "C") = Trim(MBM!MBMAddress2 & " " & MBM!MBMAddress3)
                objWorksheet.Cells(12, "C") = Trim(MBM!MBMPostCode & " " & Trim(MBM!CTYCode) & " " & Trim(MBM!STACode))
                objWorksheet.Cells(8, "I") = PVRec!PVNo
                objWorksheet.Cells(9, "I") = PVRec!PVChequeNo
                objWorksheet.Cells(10, "I") = PVRec!PVDate
                objWorksheet.Cells(37, "C") = Logon
                
                If Trim(TxtPVNoFr) = Trim(TxtPVNoTo) And Len(PVRec!Remark) <> 0 Then
                    objWorksheet.Range("A13" + ":" + "J13").RowHeight = 30
                    objWorksheet.Cells(13, "A") = "Remark : " + Trim(PVRec!Remark)
                End If
                
                MBM.Close
                Set MBM = Nothing
                
                ReceiptNo = ""
                RecSql = "select ReceiptNo from ReceiptfrMNRB where CASNumber = '" & PVRec!CasNumber & "'"
                REC.Open RecSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
                    If REC.EOF = False Then
                        ReceiptNo = REC!ReceiptNo
                    End If
                REC.Close
                Set REC = Nothing
                tmpBankName = IIf(Len(PVRec!MBMBankName), PVRec!MBMBankName, "")
                tmpBankACCNo = IIf(Len(PVRec!MBMBankACNo), PVRec!MBMBankACNo, "")
                
                If PVRec!MBMEmployeeNo <> "" Then
                    objWorksheet.Cells(31 + CDbl(i), "A") = "Employee Number : " & " " & PVRec!MBMEmployeeNo
                End If

                If Len(tmpBankACCNo) Then
                    objWorksheet.Cells(32 + CDbl(i), "A") = "Bank : " & " " & tmpBankName
                    objWorksheet.Cells(33 + CDbl(i), "A") = "Account No : " & " " & tmpBankACCNo
                End If
                
                End If
                
                PatientName = tmpMBMName
                If PVRec!coverid <> "00" Then
                    PatientSql = "Select MBMCName as PatientName from MBMCoveredPersons where MBMCNumber='" & PVRec!MBMNumber & "' AND MBMCCoverID ='" & PVRec!coverid & "'"
                
                    PatientDB.Open PatientSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    
                    If Len(PatientDB!PatientName) Then
                        PatientName = Trim(PatientDB!PatientName)
                    End If
                    PatientDB.Close
                    Set PatientDB = Nothing
                End If
                
                If PVRec!coverid = "00" Then
                PatientSql = "Select MBMName as PatientName from MBM where MBMNumber='" & PVRec!MBMNumber & "'"
                PatientDB.Open PatientSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    
                    If Len(PatientDB!PatientName) Then
                        PatientName = Trim(PatientDB!PatientName)
                    End If
                    PatientDB.Close
                    Set PatientDB = Nothing
                End If
                
                
                totAmtPaid = IIf(IsNumeric(PVRec!MBMPaidAmt), PVRec!MBMPaidAmt, 0)
                totPostage = CDbl(totPostage) + IIf(IsNumeric(PVRec!CLMPostageFee), PVRec!CLMPostageFee, 0)
                totPayeble2M = CDbl(totPayeble2M) + IIf(IsNumeric(PVRec!CLMPayableByMedix2M), PVRec!CLMPayableByMedix2M, 0)
                objWorksheet.Cells(17 + CDbl(i), "A") = CDbl(i) + 1
                objWorksheet.Cells(17 + CDbl(i), "B") = PVRec!CasNumber
                objWorksheet.Cells(17 + CDbl(i), "C") = PVRec!InvoiceNo
                objWorksheet.Cells(17 + CDbl(i), "D") = PVRec!DateVisit
                'objWorksheet.Cells(15 + CDbl(i), "E") = ReceiptNo
                objWorksheet.Cells(17 + CDbl(i), "F") = PatientName
                objWorksheet.Cells(17 + CDbl(i), "H") = PVRec!CLMTotalActual
                objWorksheet.Cells(17 + CDbl(i), "I") = 0
                objWorksheet.Cells(17 + CDbl(i), "J") = Format(CDbl(PVRec!CLMPayableByMedix2M) - CDbl(totPostage), "#,##0.00")
    
    
                i = CDbl(i) + 1
            PVRec.MoveNext
        Loop
        tmppvnorpdf = tmpPVNo
        TmpPVCode = Mid(tmpPVNo, 1, 2)
        tmpPVNo = Mid(tmpPVNo, 3, 9) + 1
        tmpPVNo = TmpPVCode & tmpPVNo
        If i > 0 Then
            objWorksheet.Cells(34, "J") = CDbl(totPayeble2M) - CDbl(totPostage)
            Dim fPDF As String
            If PrintPDFstatus = True Then
                 X = CreatePath(DocPath & "ClinicPVS\" & Trim(tmppvnorpdf))
                DocSavePath = DocPath & "ClinicPVS\" & Trim(tmppvnorpdf) & "\"
                fPDF = DocPath & "ClinicPVS\" & Trim(tmppvnorpdf) & "\" & Trim(tmppvnorpdf) & ".pdf"
                objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fPDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ElseIf PrintPDFstatus = False Then
                objWorkbook.PrintOut
                fPDF = "C:\Template\LatestPV.xls"
                objWorkbook.SaveAs fPDF
            End If
            'objWorkbook.SaveAs (LocCardPath & "PaymentVoucher\" & "tmpPaycheque" & ".xlt")
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            Set objExcel = Nothing
        End If
        totPostage = 0
        totPayeble2M = 0

        PVRec.Close
        Set PVRec = Nothing
        
    Loop

ExitFunction:
            txtPrintingStatus = "Printing Not In Progress!"
            CmdPrint.Enabled = False
            CmdExit.Enabled = True
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            txtPrintingStatus = "Printing Not In Progress!"
            CmdPrint.Enabled = True
            CmdExit.Enabled = True
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub PVPrintingCLN()
Const METHOD_NAME = "PVPrintingCLN"
On Error GoTo debugs

'On Error Resume Next
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim strsqlcnt As String
Dim RecordFound As Boolean
Dim PVRec As New ADODB.Recordset
Dim PVRec1 As New ADODB.Recordset
Dim RecFound As Boolean
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Cnt As Integer
Dim tmpPVCNTFr As String
Dim tmpPVCNTTo As String
Dim TmpPVCode As String
Dim tmpPVCnt As Integer

Dim tmpCOMPV As String


tmpNewRow = 0
tmpTotalBillAmt = 0
tmpTotalFFSAmt = 0
tmpTotalAmt = 0
tmpTotalCopayment = 0
tmpTotalClaimAmt = 0
tmpTotalClaimAmtSub = 0


Dim strsqlGrp As String
    strAppend = ""
    tmpPVCNTFr = Mid(TxtPVNoFr, 3, 9)
    tmpPVCNTTo = Mid(TxtPVNoTo, 3, 9)
    TmpPVCode = Mid(TxtPVNoFr, 1, 2)
    
    tmpPVCnt = 1
    RecordFound = False
    Do While tmpPVCNTFr <= tmpPVCNTTo
        tmpCOMPV = TmpPVCode & tmpPVCNTFr
        tmpTotalClaimAmtSub = 0
        tmpTotalBillAmt = 0
        tmpTotalFFSAmt = 0
        tmpTotalAmt = 0
        tmpTotalCopayment = 0
        tmpTotalClaimAmt = 0
        StrSql = ""
        StrSql = " SELECT CISDB.dbo.CISCas.ClinicCode, dbo.Payment2CLN.PVPayCLNStatus,dbo.Payment2CLN.CLNAccPeriodMth, dbo.Payment2CLN.CLNAccPeriodYr, HISDB.dbo.CLN.CLNName, " & _
                " CISWorksheet.dbo.ClaimInvoice.INVPayTo,HISDB.dbo.CLN.CLNAdd1, HISDB.dbo.CLN.CLNAdd2, HISDB.dbo.CLN.CLNAdd3, " & _
                " HISDB.dbo.CLN.CLNState, HISDB.dbo.CLN.CLNCity, HISDB.dbo.CLN.CLNPostCode, dbo.Payment2CLN.PVCLNNo,dbo.Payment2CLN.PVCLNDate, dbo.Payment2CLN.PVCLNChequeDate, dbo.Payment2CLN.PVCLNChequeNo, CISDB.dbo.CISCas.InvoiceNo, " & _
                " CISDB.dbo.CISCas.InvoiceDate, dbo.Payment2CLN.PVCLNChequeAmt, CISDB.dbo.CISCas.CASStatus, dbo.Payment2CLN.PVCLNBank,dbo.Payment2CLN.PVCLNAccount,  dbo.Payment2CLN.CLNRemark, HISDB.dbo.CLN.ClinicGrpCode, " & _
                " HISDB.dbo.CLN.CLNPName, HISDB.dbo.CLN.CLNPAdd1,HISDB.dbo.CLN.CLNPAdd2, HISDB.dbo.CLN.CLNPAdd3, HISDB.dbo.CLN.CLNEmail, HISDB.dbo.CLN.CLNBankAcc, dbo.Payment2CLN.PVWordingOne,dbo.Payment2CLN.PVWordingTwo," & _
                " HISDB.dbo.CLN.CLNType, HISDB.dbo.CLN.CLNFFSStatus, HISDB.dbo.CLN.CLNPayee, HISDB.dbo.CLN.CLNRegNo, HISDB.dbo.CLN.CLNBankName, SUM(CISWorksheet.dbo.Claim.CLMTotalActual) AS CLMTotalActual,sum(CISWorksheet.dbo.Claim.CLMCoPayCLNColl) as Copay, " & _
                " sum(CISWorksheet.dbo.Claim.CLMTotalApproved ) as TotalApproved,SUM(CISWorksheet.dbo.Claim.CLMMCOProcFee) AS CLMMCOProcFee, SUM(CISWorksheet.dbo.Claim.CLMBordxAmt ) AS CLMBordxAmt,SUM(HISDB.dbo.CLN.CLNMCOProcfeePerc) AS CLNMCOProcfeePerc," & _
                " SUM(CISWorksheet.dbo.Claim.CLMPayableByMedix2C) AS CLMPayableByMedix2C,SUM(dbo.Payment2CLN.CLNPaidAmt) AS CLNPaidAmt,SUM(CISWorksheet.dbo.Claim.CLMNonGPFee) AS CLMNonGPFee ,SUM(CISWorksheet.dbo.Claim.CLMFinalApprove) AS CLMFinalApprove" & _
                " FROM HISDB.dbo.CLN INNER JOIN " & _
                " CISDB.dbo.CISCas ON HISDB.dbo.CLN.CLNCode = CISDB.dbo.CISCas.ClinicCode RIGHT OUTER JOIN   " & _
                " dbo.Payment2CLN RIGHT OUTER JOIN   " & _
                " CISWorksheet.dbo.Bordx RIGHT OUTER JOIN   " & _
                " CISWorksheet.dbo.ClaimInvoice INNER JOIN   " & _
                " CISWorksheet.dbo.Claim ON CISWorksheet.dbo.ClaimInvoice.CASNumber = CISWorksheet.dbo.Claim.CASNumber ON   " & _
                " CISWorksheet.dbo.Bordx.BordxRefNo = CISWorksheet.dbo.Claim.BordxRefNo ON dbo.Payment2CLN.CASNumber = CISWorksheet.dbo.Claim.CASNumber ON   " & _
                " CISDB.dbo.CISCAS.CasNumber = CISWorksheet.dbo.Claim.CasNumber  where CISDB.dbo.CISCas.InvoiceNo <>N'' and  CISWorksheet.dbo.ClaimInvoice.INVPayTo = 'C' And CLMFinalApprove is not null AND dbo.Payment2CLN.PVCLNNo = '" & Trim(tmpCOMPV) & "' And Claimability <> 'C'"
                    
                strsqlGrp = "GROUP BY CISDB.dbo.CISCas.ClinicCode, dbo.Payment2CLN.PVPayCLNStatus, dbo.Payment2CLN.CLNAccPeriodMth, dbo.Payment2CLN.CLNAccPeriodYr, " & _
                " HISDB.dbo.CLN.CLNName, CISWorksheet.dbo.ClaimInvoice.INVPayTo,  HISDB.dbo.CLN.CLNAdd1,   " & _
                " HISDB.dbo.CLN.CLNAdd2, HISDB.dbo.CLN.CLNAdd3, HISDB.dbo.CLN.CLNState, HISDB.dbo.CLN.CLNCity, HISDB.dbo.CLN.CLNPostCode, dbo.Payment2CLN.PVCLNNo, " & _
                " dbo.Payment2CLN.PVCLNDate, dbo.Payment2CLN.PVCLNChequeDate, dbo.Payment2CLN.PVCLNChequeNo, CISDB.dbo.CISCas.InvoiceNo, " & _
                " CISDB.dbo.CISCas.InvoiceDate, dbo.Payment2CLN.PVCLNChequeAmt, CISDB.dbo.CISCas.CASStatus, dbo.Payment2CLN.PVCLNBank, " & _
                " dbo.Payment2CLN.PVCLNAccount, dbo.Payment2CLN.CLNRemark, HISDB.dbo.CLN.ClinicGrpCode, HISDB.dbo.CLN.CLNPName, HISDB.dbo.CLN.CLNPAdd1, " & _
                " HISDB.dbo.CLN.CLNPAdd2, HISDB.dbo.CLN.CLNPAdd3, HISDB.dbo.CLN.CLNEmail, HISDB.dbo.CLN.CLNBankAcc, dbo.Payment2CLN.PVWordingOne, " & _
                " dbo.Payment2CLN.PVWordingTwo, HISDB.dbo.CLN.CLNType, " & _
                " HISDB.dbo.CLN.CLNFFSStatus , HISDB.dbo.CLN.CLNPayee, HISDB.dbo.CLN.CLNRegNo, HISDB.dbo.CLN.CLNBankName "
                    
                StrSql = StrSql & strsqlGrp
        'StrSql = ""
        'StrSql = "Select * from VIEW_PymtCLN_PV where PVCLNNo = '" & Trim(tmpCOMPV) & "'"
        PVRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        strsqlcnt = ""
        strsqlcnt = " SELECT Count(*) AS CNT_Rec FROM HISDB.dbo.CLN INNER JOIN CISDB.dbo.CISCas ON HISDB.dbo.CLN.CLNCode = CISDB.dbo.CISCas.ClinicCode RIGHT OUTER JOIN   dbo.Payment2CLN RIGHT OUTER JOIN   CISWorksheet.dbo.Bordx RIGHT OUTER JOIN   CISWorksheet.dbo.ClaimInvoice INNER JOIN    " & _
                    " CISWorksheet.dbo.Claim ON CISWorksheet.dbo.ClaimInvoice.CASNumber = CISWorksheet.dbo.Claim.CASNumber ON   CISWorksheet.dbo.Bordx.BordxRefNo = CISWorksheet.dbo.Claim.BordxRefNo ON dbo.Payment2CLN.CASNumber = CISWorksheet.dbo.Claim.CASNumber ON   " & _
                    " CISDB.dbo.CISCAS.CasNumber = CISWorksheet.dbo.Claim.CasNumber  where CISDB.dbo.CISCas.InvoiceNo <>N'' and  CISWorksheet.dbo.ClaimInvoice.INVPayTo = 'C' And CLMFinalApprove is not null AND dbo.Payment2CLN.PVCLNNo = '" & Trim(tmpCOMPV) & "' And Claimability <> 'C'"
        PVRec1.Open strsqlcnt, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        
        tmp_Rec = 0
        tmp_Rec = PVRec1!CNT_Rec
        
        
        RecordFound = False
        Cnt = 0
        tmpTotalFFSAmt = 0
        
        Do While Not PVRec.EOF
            RecordFound = True
            If objExcel Is Nothing Then
                Set objExcel = CreateObject("Excel.application")
                    objExcel.Interactive = False
                    objExcel.Visible = False
                objExcel.WindowState = xlMaximized
                Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-ClinicTest.xlt")
                Set objWorksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objWorksheet = objWorkbook.Worksheets.Item(1)
            End If
                
            'On Error GoTo HANDLE_ERROR
                'Cnt = 17
                Cnt = CDbl(Cnt) + 1
                If Cnt = 1 Then
                    objWorksheet.Cells(8, "C") = IIf(Len(PVRec!CLNPName), PVRec!CLNPName, "")
                    objWorksheet.Cells(9, "C") = PVRec!CLNName & " (" & PVRec!ClinicCode & ")"
                    objWorksheet.Cells(10, "C") = PVRec!CLNPAdd1
                    objWorksheet.Cells(11, "C") = PVRec!CLNPAdd2
                    objWorksheet.Cells(12, "C") = PVRec!CLNPAdd3
                    objWorksheet.Cells(8, "H") = PVRec!PVCLNNo '& "/" & Format(PVRec!PVCLNDate, "MM/YY")
                    objWorksheet.Cells(9, "H") = PVRec!PVCLNChequeNo
                    objWorksheet.Cells(10, "H") = PVRec!PVCLNDate
                    tmpNewRow = 17 + CDbl(Cnt)
                End If
                
                
                ''
              
                If CDbl(tmpNewRow) <= "44" Then
                    
                    objWorksheet.Cells(16 + CDbl(Cnt), "A") = Cnt
                    objWorksheet.Cells(16 + CDbl(Cnt), "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                    objWorksheet.Cells(16 + CDbl(Cnt), "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "D") = PVRec!InvoiceDate
                    objWorksheet.Cells(16 + CDbl(Cnt), "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "F") = PVRec!CLMTotalActual
                    objWorksheet.Cells(16 + CDbl(Cnt), "F").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    objWorksheet.Cells(16 + CDbl(Cnt), "G").Borders.LineStyle = XlLineStyle.xlContinuous
                    'objWorksheet.Cells(16 + CDbl(Cnt), "H") = IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                    objWorksheet.Cells(16 + CDbl(Cnt), "H") = CDbl(PVRec!CLMTotalActual) - CDbl(PVRec!copay)
                    tmpTotalClaimAmtSub = 0
                    tmpTotalClaimAmtSub = CDbl(PVRec!CLMTotalActual) - CDbl(PVRec!copay)
                    objWorksheet.Cells(16 + CDbl(Cnt), "H").Borders.LineStyle = XlLineStyle.xlContinuous
                    
                    objWorksheet.Cells(16 + CDbl(Cnt), "I") = IIf(Len(PVRec!CLMMCOProcFee), PVRec!CLMMCOProcFee, 0)
                    objWorksheet.Cells(16 + CDbl(Cnt), "I").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "J") = IIf(Len(PVRec!CLMPayableByMedix2C), CDbl(PVRec!CLMPayableByMedix2C), 0)
                    objWorksheet.Cells(16 + CDbl(Cnt), "J").Borders.LineStyle = XlLineStyle.xlContinuous
                                        
                    Dim Fnlapp As Double
                    If IsNull(PVRec!CLMTotalActual) Then
                        Fnlapp = 0
                    Else
                        Fnlapp = PVRec!CLMTotalActual
                    End If
                    
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + Fnlapp
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = CDbl(tmpTotalClaimAmt) + CDbl(tmpTotalClaimAmtSub)
                    tmpNewRow = CDbl(tmpNewRow) + 1
                
                ElseIf CDbl(tmpNewRow) > "44" And CDbl(tmpNewRow) <= "82" Then
                    
                    'Cnt = CDbl(Cnt) + 14
                    If tmpNewRow = "45" Then
                        

                        'Cnt = CDbl(tmpNewRow) + 11
                        'objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "PARTICULARS"
                        'objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = "RM"
                        'tmpNewRow = CDbl(tmpNewRow) + 1
                        
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        
                    
                    End If
                    
                   
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                    
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                    objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                                                           
                    
                    'tmpTotalBillAmt = IIf(Len((tmpTotalBillAmt)), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    'tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    'tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len((PVRec!CLNPaidAmt)), CDbl(PVRec!CLNPaidAmt), 0)
                    
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len(PVRec!CLMTotalActual), PVRec!CLMTotalActual, 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                    
                        
                    tmpNewRow = CDbl(tmpNewRow) + 1
                ElseIf CDbl(tmpNewRow) > "82" And CDbl(tmpNewRow) <= "133" Then
                    If tmpNewRow = "83" Then
                        
                        tmpNewRow = CDbl(tmpNewRow) + 14
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        
                    
                    End If
                    
                    
                    'tmpNewRow = CDbl(tmpNewRow) + 1
                    
                    
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                                                           
                        tmpNewRow = CDbl(tmpNewRow) + 1
                    'tmpTotalBillAmt = CDbl(tmpTotalBillAmt) + CDbl(PVRec!CLMFinalApprove)
                    'tmpTotalFFSAmt = CDbl(tmpTotalFFSAmt) + CDbl(PVRec!CLMMCOProcFee)
                    'tmpTotalAmt = CDbl(tmpTotalAmt) + CDbl(PVRec!CLNPaidAmt)
                    
                   ' tmpTotalBillAmt = IIf(Len((tmpTotalBillAmt)), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                   ' tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                   ' tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len((PVRec!CLNPaidAmt)), CDbl(PVRec!CLNPaidAmt), 0)
                        
                        
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                        
                ElseIf CDbl(tmpNewRow) > "133" And CDbl(tmpNewRow) <= "184" Then
                    If tmpNewRow = "134" Then
                        
                        tmpNewRow = CDbl(tmpNewRow) + 14
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        
                    
                    End If
                    
                    
                    'tmpNewRow = CDbl(tmpNewRow) + 1
                    
                    
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                                                           
                        tmpNewRow = CDbl(tmpNewRow) + 1
                    
                        'tmpTotalBillAmt = CDbl(tmpTotalBillAmt) + CDbl(PVRec!CLMFinalApprove)
                        'tmpTotalFFSAmt = CDbl(tmpTotalFFSAmt) + CDbl(PVRec!CLMMCOProcFee)
                        'tmpTotalAmt = CDbl(tmpTotalAmt) + CDbl(PVRec!CLNPaidAmt)
                        
                    'tmpTotalBillAmt = IIf(Len((tmpTotalBillAmt)), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    'tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    'tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len((PVRec!CLNPaidAmt)), CDbl(PVRec!CLNPaidAmt), 0)
                        
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                    
                ElseIf CDbl(tmpNewRow) > "184" And CDbl(tmpNewRow) <= "235" Then
                    If tmpNewRow = "185" Then
                        
                        tmpNewRow = CDbl(tmpNewRow) + 14
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        
                    End If
                    
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                                                           
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        'tmpTotalBillAmt = IIf(Len((tmpTotalBillAmt)), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                        'tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                        'tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len((PVRec!CLNPaidAmt)), CDbl(PVRec!CLNPaidAmt), 0)
                        
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                                        
                ElseIf CDbl(tmpNewRow) > "235" And CDbl(tmpNewRow) <= "286" Then
                    If tmpNewRow = "236" Then
                        
                        tmpNewRow = CDbl(tmpNewRow) + 14
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        
                    
                    End If
                    
                    
                    'tmpNewRow = CDbl(tmpNewRow) + 1
                    
                    
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                        
                        tmpNewRow = CDbl(tmpNewRow) + 1
                                            
                        'tmpTotalBillAmt = IIf(Len((tmpTotalBillAmt)), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                        'tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                        'tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len((PVRec!CLNPaidAmt)), CDbl(PVRec!CLNPaidAmt), 0)
                        
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                    
                ElseIf CDbl(tmpNewRow) > "286" And CDbl(tmpNewRow) <= "338" Then
                    If tmpNewRow = "287" Then
                        
                        tmpNewRow = CDbl(tmpNewRow) + 14
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                    
                    End If
                    
                    'tmpNewRow = CDbl(tmpNewRow) + 1
                    
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        'tmpTotalBillAmt = CDbl(tmpTotalBillAmt) + CDbl(PVRec!CLMFinalApprove)
                        'tmpTotalFFSAmt = CDbl(tmpTotalFFSAmt) + CDbl(PVRec!CLMMCOProcFee)
                        'tmpTotalAmt = CDbl(tmpTotalAmt) + CDbl(PVRec!CLNPaidAmt)
                        'tmpTotalBillAmt = IIf(Len((tmpTotalBillAmt)), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                        'tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                        'tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len((PVRec!CLNPaidAmt)), CDbl(PVRec!CLNPaidAmt), 0)
                        
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                        
                'extra code 21-07-2016
                ElseIf CDbl(tmpNewRow) > "338" And CDbl(tmpNewRow) <= "389" Then
                    If tmpNewRow = "339" Then
                        
                        tmpNewRow = CDbl(tmpNewRow) + 14
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                    End If
                    'tmpNewRow = CDbl(tmpNewRow) + 1
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                                            
                        'tmpTotalBillAmt = IIf(Len((tmpTotalBillAmt)), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                        'tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                        'tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len((PVRec!CLNPaidAmt)), CDbl(PVRec!CLNPaidAmt), 0)
                        
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                ElseIf CDbl(tmpNewRow) > "389" And CDbl(tmpNewRow) <= "440" Then
                    If tmpNewRow = "390" Then
                        tmpNewRow = CDbl(tmpNewRow) + 14
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                    End If
                    'tmpNewRow = CDbl(tmpNewRow) + 1
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                ElseIf CDbl(tmpNewRow) > "440" And CDbl(tmpNewRow) <= "491" Then
                    If tmpNewRow = "441" Then
                        tmpNewRow = CDbl(tmpNewRow) + 14
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                    End If
                    'tmpNewRow = CDbl(tmpNewRow) + 1
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A") = Cnt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D") = PVRec!InvoiceDate
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = PVRec!CLMTotalActual
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = PVRec!CLMBordxAmt
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = PVRec!CLMMCOProcFee
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = CDbl(PVRec!CLMPayableByMedix2C)
                        objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len((PVRec!CLMFinalApprove)), CDbl(PVRec!CLMFinalApprove), 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                ElseIf CDbl(tmpNewRow) > "491" And CDbl(tmpNewRow) <= "541" Then
                
                Dim Tamptotalcnt As Long
                Tamptotalcnt = tmp_Rec
                
                If tmpNewRow >= "491" Then
                
                 'Call BindExcelvlaues
                 
                 End If
            
            End If
            PVRec.MoveNext
        Loop
        'cmd.Parameters.Refresh
            If CDbl(tmpNewRow) <= "44" Then
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "F") = tmpTotalBillAmt
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "G") = tmpTotalCopayment
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "H") = tmpTotalClaimAmt
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "I") = tmpTotalFFSAmt
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "J") = tmpTotalAmt
                objWorksheet.Cells(CDbl(tmpNewRow) - 1, "J").Borders.LineStyle = XlLineStyle.xlContinuous
            Else
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = tmpTotalBillAmt
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = tmpTotalCopayment
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = tmpTotalClaimAmt
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I") = tmpTotalFFSAmt
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J") = tmpTotalAmt
                objWorksheet.Cells(CDbl(tmpNewRow) + 11, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                
                
            End If
            
            If CDbl(tmpNewRow) <= "43" Then
                objWorksheet.Cells(46, "C") = "Prepared By : " & Logon
                'objWorksheet.Cells(46, "C") = Logon
                objWorksheet.Cells(46, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(46, "E") = "______________"
                'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(48, "C") = "Approved By : "
                objWorksheet.Cells(48, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(48, "E") = "______________"
                objWorksheet.Cells(50, "C") = "Signatories : "
                objWorksheet.Cells(50, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(50, "E") = "______________"
                objWorksheet.Cells(46, "G") = "Acknowledgement of receipt : "
                objWorksheet.Cells(49, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(49, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(50, "G") = "Signature & Company Stamp"
                
             '   objWorksheet.Cells(CDbl(tmpNewRow) + 11, "F") = tmpTotalBillAmt
             '   objWorksheet.Cells(CDbl(tmpNewRow) + 11, "G") = tmpTotalFFSAmt
             '   objWorksheet.Cells(CDbl(tmpNewRow) + 11, "H") = tmpTotalAmt

            ElseIf CDbl(tmpNewRow) > 43 And CDbl(tmpNewRow) <= 82 Then
                
                objWorksheet.Cells(46, "C") = "Prepared By : " & Logon
                'objWorksheet.Cells(46, "C") = Logon
                objWorksheet.Cells(46, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(46, "E") = "______________"
                'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(48, "C") = "Approved By : "
                objWorksheet.Cells(48, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(48, "E") = "______________"
                objWorksheet.Cells(50, "C") = "Signatories : "
                objWorksheet.Cells(50, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(50, "E") = "______________"
                objWorksheet.Cells(46, "G") = "Acknowledgement of receipt : "
                objWorksheet.Cells(49, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(49, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(50, "G") = "Signature & Company Stamp"
                
                
                objWorksheet.Cells(97, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(97, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(97, "E") = "______________"
                'objWorksheet.Cells(97, "D") = Logon
                objWorksheet.Cells(99, "C") = "Approved By : "
                objWorksheet.Cells(99, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(99, "E") = "______________"
                objWorksheet.Cells(101, "C") = "Signatories : "
                objWorksheet.Cells(101, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(101, "E") = "______________"
                objWorksheet.Cells(97, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(100, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(100, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(101, "G") = "Signature & Company Stamp"
                
                
            ElseIf CDbl(tmpNewRow) > 82 And CDbl(tmpNewRow) <= 133 Then
                objWorksheet.Cells(46, "C") = "Prepared By : " & Logon
                'objWorksheet.Cells(46, "C") = Logon
                objWorksheet.Cells(46, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(46, "E") = "______________"
                'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(48, "C") = "Approved By : "
                objWorksheet.Cells(48, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(48, "E") = "______________"
                objWorksheet.Cells(50, "C") = "Signatories : "
                objWorksheet.Cells(50, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(50, "E") = "______________"
                objWorksheet.Cells(46, "G") = "Acknowledgement of receipt : "
                objWorksheet.Cells(49, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(49, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(50, "G") = "Signature & Company Stamp"
                
                
                objWorksheet.Cells(97, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(97, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(97, "E") = "______________"
                'objWorksheet.Cells(97, "D") = Logon
                objWorksheet.Cells(99, "C") = "Approved By : "
                objWorksheet.Cells(99, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(99, "E") = "______________"
                objWorksheet.Cells(101, "C") = "Signatories : "
                objWorksheet.Cells(101, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(101, "E") = "______________"
                objWorksheet.Cells(97, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(100, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(100, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(101, "G") = "Signature & Company Stamp"
                
                objWorksheet.Cells(148, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(148, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(148, "E") = "______________"
                'objWorksheet.Cells(148, "D") = Logon
                objWorksheet.Cells(150, "C") = "Approved By : "
                objWorksheet.Cells(150, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(150, "E") = "______________"
                objWorksheet.Cells(152, "C") = "Signatories : "
                objWorksheet.Cells(152, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(152, "E") = "______________"
                objWorksheet.Cells(148, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(151, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(151, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(152, "G") = "Signature & Company Stamp"
                
            ElseIf CDbl(tmpNewRow) > "133" And CDbl(tmpNewRow) <= "184" Then
                objWorksheet.Cells(46, "C") = "Prepared By : " & Logon
                'objWorksheet.Cells(46, "C") = Logon
                objWorksheet.Cells(46, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(46, "E") = "______________"
                'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(48, "C") = "Approved By : "
                objWorksheet.Cells(48, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(48, "E") = "______________"
                objWorksheet.Cells(50, "C") = "Signatories : "
                objWorksheet.Cells(50, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(50, "E") = "______________"
                objWorksheet.Cells(46, "G") = "Acknowledgement of receipt : "
                objWorksheet.Cells(49, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(49, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(50, "G") = "Signature & Company Stamp"
                
                
                objWorksheet.Cells(97, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(97, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(97, "E") = "______________"
                'objWorksheet.Cells(97, "D") = Logon
                objWorksheet.Cells(99, "C") = "Approved By : "
                objWorksheet.Cells(99, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(99, "E") = "______________"
                objWorksheet.Cells(101, "C") = "Signatories : "
                objWorksheet.Cells(101, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(101, "E") = "______________"
                objWorksheet.Cells(97, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(100, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(100, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(101, "G") = "Signature & Company Stamp"
                
                objWorksheet.Cells(148, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(148, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(148, "E") = "______________"
                'objWorksheet.Cells(148, "D") = Logon
                objWorksheet.Cells(150, "C") = "Approved By : "
                objWorksheet.Cells(150, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(150, "E") = "______________"
                objWorksheet.Cells(152, "C") = "Signatories : "
                objWorksheet.Cells(152, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(152, "E") = "______________"
                objWorksheet.Cells(148, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(151, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(151, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(152, "G") = "Signature & Company Stamp"
                
                objWorksheet.Cells(199, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(199, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(148, "E") = "______________"
                'objWorksheet.Cells(199, "D") = Logon
                objWorksheet.Cells(201, "C") = "Approved By : "
                objWorksheet.Cells(201, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(150, "E") = "______________"
                objWorksheet.Cells(203, "C") = "Signatories : "
                objWorksheet.Cells(203, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(152, "E") = "______________"
                objWorksheet.Cells(199, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(202, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(202, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(203, "G") = "Signature & Company Stamp"
                
            ElseIf CDbl(tmpNewRow) > "184" And CDbl(tmpNewRow) <= "235" Then
                objWorksheet.Cells(46, "C") = "Prepared By : " & Logon
                'objWorksheet.Cells(46, "C") = Logon
                objWorksheet.Cells(46, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(46, "E") = "______________"
                'objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(48, "C") = "Approved By : "
                objWorksheet.Cells(48, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(48, "E") = "______________"
                objWorksheet.Cells(50, "C") = "Signatories : "
                objWorksheet.Cells(50, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(50, "E") = "______________"
                objWorksheet.Cells(46, "G") = "Acknowledgement of receipt : "
                objWorksheet.Cells(49, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(49, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(50, "G") = "Signature & Company Stamp"
                
                
                objWorksheet.Cells(97, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(97, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(97, "E") = "______________"
                'objWorksheet.Cells(97, "D") = Logon
                objWorksheet.Cells(99, "C") = "Approved By : "
                objWorksheet.Cells(99, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(99, "E") = "______________"
                objWorksheet.Cells(101, "C") = "Signatories : "
                objWorksheet.Cells(101, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(101, "E") = "______________"
                objWorksheet.Cells(97, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(100, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(100, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(101, "G") = "Signature & Company Stamp"
                
                objWorksheet.Cells(148, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(148, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(148, "E") = "______________"
                'objWorksheet.Cells(148, "D") = Logon
                objWorksheet.Cells(150, "C") = "Approved By : "
                objWorksheet.Cells(150, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(150, "E") = "______________"
                objWorksheet.Cells(152, "C") = "Signatories : "
                objWorksheet.Cells(152, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(152, "E") = "______________"
                objWorksheet.Cells(148, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(151, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(151, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(152, "G") = "Signature & Company Stamp"
                
                objWorksheet.Cells(199, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(199, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(148, "E") = "______________"
                'objWorksheet.Cells(199, "D") = Logon
                objWorksheet.Cells(201, "C") = "Approved By : "
                objWorksheet.Cells(201, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(150, "E") = "______________"
                objWorksheet.Cells(203, "C") = "Signatories : "
                objWorksheet.Cells(203, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(152, "E") = "______________"
                objWorksheet.Cells(199, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(202, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(202, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(203, "G") = "Signature & Company Stamp"
                
                objWorksheet.Cells(250, "C") = "Prepared By : " & Logon
                objWorksheet.Cells(250, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(148, "E") = "______________"
                'objWorksheet.Cells(250, "D") = Logon
                objWorksheet.Cells(252, "C") = "Approved By : "
                objWorksheet.Cells(252, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(150, "E") = "______________"
                objWorksheet.Cells(254, "C") = "Signatories : "
                objWorksheet.Cells(254, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                'objWorksheet.Cells(152, "E") = "______________"
                objWorksheet.Cells(250, "G") = "Acknowledgement of receipt:"
                objWorksheet.Cells(253, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(253, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                objWorksheet.Cells(254, "G") = "Signature & Company Stamp"
                
            End If
        PVRec.Close
        Set PVRec = Nothing
        If RecordFound = True Then
            'objWorkbook.PrintOut
            
            objWorkbook.SaveAs (LocCardPath & "PaymentVoucher\" & "tmpPaycheque" & ".xlt")
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            Set objExcel = Nothing
        Else
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            Set objExcel = Nothing

        End If
        
        tmpPVCNTFr = CDbl(tmpPVCNTFr) + 1
    Loop
ExitFunction:
            txtPrintingStatus = "Printing Not In Progress!"
            CmdPrint.Enabled = False
            CmdExit.Enabled = True
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            txtPrintingStatus = "Printing Not In Progress!"
            CmdPrint.Enabled = True
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            CmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdprintPDF_Click()
Const METHOD_NAME = "cmdprintPDF_Click"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
    PrintPDFstatus = True
    Screen.MousePointer = vbHourglass
    If (lsPVMBMListing.ListItems.Count > 0 And chqMode = "MedixPayMBM") Or (lsPVListing.ListItems.Count > 0 And chqMode = "MedixPayCLN") Then
        If MsgBox("Do you want to print PDF now?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
            RecFound = False
            txtPrintingStatus = "Printing In Progress!"
            CmdPrint.Enabled = False
            CmdExit.Enabled = False
            
                If chqMode = "MedixPayMBM" Then
                    Call PVPrintingMBM
                ElseIf chqMode = "MedixPayCLN" Then
                    Call PVPrintingCLN_New
                End If
        End If
    Else
        MsgBox "Please select PV to print", vbOKOnly + vbCritical
    End If
    Screen.MousePointer = vbDefault
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub cmdSearch_Click()
Dim RecSel As Double
    RecSel = 0
    lblTotalchqSel = 0
    If TxtPVNoFr = "" Then
        MsgBox "Please Enter PV No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If TxtPVNoTo = "" Then
        MsgBox "Please Enter PV No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lblTotalchqSel = RecSel
    lsPVListing.ListItems.Clear
    lsPVMBMListing.ListItems.Clear
       
           If chqMode = "MedixPayMBM" Then
                Call PVMBMListing
                lsPVListing.Visible = False
                lsPVMBMListing.Visible = True
            ElseIf chqMode = "MedixPayCLN" Then
                Call PVCLNListing
                lsPVListing.Visible = True
                lsPVMBMListing.Visible = False
            End If
    End Sub

Private Sub PVCLNListing()
Const METHOD_NAME = "PVCLNListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Record As ListItem
Dim RecordFound As Boolean
Dim PVRec As New ADODB.Recordset
Dim StrCLNSql As String
Dim ClnRec As New ADODB.Recordset
Dim tmpPVCLNBank
Dim tmpPVCLNAccount As String
Dim tmpCLNRemark As String
Dim tmpCLNBankAcc As String
Dim tmpCLNPName As String
Dim tmpCLNPAdd1 As String
Dim tmpCLNPAdd2 As String
Dim tmpCLNPAdd3 As String
Dim tmpCLNType As String
Dim tmpCLNFFSStatus As String
Dim tmpCLNRegNo As String
Dim tmpCLNMCOProcfeePerc As String
Dim tmpCLNBankName As String
Dim tmpCLNName As String
Dim tmpCLNAdd1  As String
Dim tmpCLNAdd2  As String
Dim tmpCLNAdd3 As String
Dim tmpPCode  As String
Dim tmpCity  As String
Dim tmpState   As String
    
    RecordFound = False
    strAppend = ""
    If Len(TxtPVNoFr) Then
        strAppend = strAppend & " dbo.Payment2CLN.PVCLNNo >= '" & Trim(TxtPVNoFr) & "'"
    End If
        
    If Len(TxtPVNoTo) Then
        strAppend = strAppend & " and dbo.Payment2CLN.PVCLNNo <= '" & Trim(TxtPVNoTo) & "'"
    End If
        
    strAppend = strAppend & " and InvoiceNo <> '' "
    'StrAppend = StrAppend & " Order by PVCLNNo"
    
    Set cmd.ActiveConnection = MediConnCISPayment
    cmd.CommandText = "SP_CISPayment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("PaymentType", adVarChar, adParamInput, 10, "Pay2CLN")
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm2
    PVRec.CursorLocation = adUseClient
    PVRec.CursorType = adOpenForwardOnly
    PVRec.CacheSize = 100
    Set PVRec = cmd.Execute
    
    Do While Not PVRec.EOF
        RecordFound = True
        Set Record = lsPVListing.ListItems.Add(, , PVRec!PVCLNNo)
        
        tmpCLNName = Trim(PVRec!CLNName)
        tmpCLNAdd1 = IIf(Len(PVRec!CLNAdd1), PVRec!CLNAdd1, "")
        tmpCLNAdd2 = IIf(Len(PVRec!CLNAdd2), PVRec!CLNAdd2, "")
        tmpCLNAdd3 = IIf(Len(PVRec!CLNAdd3), PVRec!CLNAdd3, "")
        tmpPCode = IIf(Len(PVRec!CLNPostCode), PVRec!CLNPostCode, "")
        tmpCity = IIf(Len(PVRec!CLNCity), PVRec!CLNCity, "")
        tmpState = IIf(Len(PVRec!CLNState), PVRec!CLNState, "")
        tmpCLNRemark = IIf(Len(PVRec!CLNRemark), PVRec!CLNRemark, "")
        tmpCLNBankAcc = IIf(Len(PVRec!CLNBankAcc), PVRec!CLNBankAcc, "")
        tmpCLNPName = IIf(Len(PVRec!CLNPName), PVRec!CLNPName, "")
        tmpCLNPAdd1 = IIf(Len(PVRec!CLNPAdd1), PVRec!CLNPAdd1, "")
        tmpCLNPAdd2 = IIf(Len(PVRec!CLNPAdd2), PVRec!CLNPAdd2, "")
        tmpCLNPAdd3 = IIf(Len(PVRec!CLNPAdd3), PVRec!CLNPAdd3, "")
        tmpCLNType = IIf(Len(PVRec!CLNType), PVRec!CLNType, "")
        tmpCLNFFSStatus = IIf(Len(PVRec!CLNFFSStatus), PVRec!CLNFFSStatus, "")
        tmpCLNPName = IIf(Len(PVRec!CLNPName), PVRec!CLNPName, "")
        tmpCLNRegNo = IIf(Len(PVRec!CLNRegNo), PVRec!CLNRegNo, "")
        tmpCLNMCOProcfeePerc = IIf(Len(PVRec!CLNMCOProcfeePerc), PVRec!CLNMCOProcfeePerc, "")
        tmpCLNBankName = IIf(Len(PVRec!CLNBankName), PVRec!CLNBankName, "")
        tmpCLNBankAcc = IIf(Len(PVRec!CLNBankAcc), PVRec!CLNBankAcc, "")
        
        If Len(PVRec!PVCLNDate) Then
            Record.SubItems(1) = (PVRec!PVCLNDate)
        End If
        If Len(tmpCLNPName) Then
            Record.SubItems(2) = (tmpCLNPName)
        End If
        If Len(PVRec!InvoiceNo) Then
            Record.SubItems(3) = (PVRec!InvoiceNo)
        End If
        If Len(PVRec!InvoiceDate) Then
            Record.SubItems(4) = (PVRec!InvoiceDate)
        End If
        If Len(PVRec!PVCLNChequeNo) Then
            Record.SubItems(5) = (PVRec!PVCLNChequeNo)
        End If
        If Len(PVRec!PVCLNDate) Then
            Record.SubItems(6) = (PVRec!PVCLNDate)
        End If
        If Len(PVRec!PVCLNChequeAmt) Then
            Record.SubItems(7) = (PVRec!PVCLNChequeAmt)
        End If
        If Len(tmpCLNName) Then
            Record.SubItems(8) = (tmpCLNName)
        End If
        If Len(PVRec!ClinicCode) Then
            Record.SubItems(9) = (PVRec!ClinicCode)
        End If
        If Len(tmpCLNPAdd1) Then
            Record.SubItems(10) = (tmpCLNPAdd1)
        End If
        If Len(tmpCLNPAdd2) Then
            Record.SubItems(11) = (tmpCLNPAdd2)
        End If
        If Len(tmpCLNPAdd3) Then
            Record.SubItems(12) = (tmpCLNPAdd3)
        End If
        'If Len(PVRec!BordxRefNo) Then
        '    Record.SubItems(13) = (PVRec!BordxRefNo)
        'End If
        If Len(PVRec!CLMFinalApprove) Then
            Record.SubItems(14) = (PVRec!CLMFinalApprove)
        End If
        If Len(PVRec!CLMMCOProcFee) Then
            Record.SubItems(15) = (PVRec!CLMMCOProcFee)
        End If
        
        'RecCnt = RecCnt + 1
        'lblTotalChqList = RecCnt
        PVRec.MoveNext
    Loop
    PVRec.Close
    Set PVRec = Nothing
    
    CmdExit.Enabled = True
    
    If RecordFound = True Then
        'If CDbl(lblTotalChqList) = CDbl(lblTotalchqSel) Then
            CmdPrint.Enabled = True
            cmdprintPDF.Enabled = True
        'Else
        '    cmdPrint.Enabled = False
        '    MsgBox "Cheque range not in series", vbOKOnly + vbCritical
        'End If
    Else
        MsgBox "No record found", vbOKOnly + vbCritical
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub PVMBMListing()

Const METHOD_NAME = "PVMBMListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Record As ListItem
Dim RecordFound As Boolean
Dim PVRec As New ADODB.Recordset
Dim sql As String

    RecordFound = False
    strAppend = ""
    sql = ""
    If Len(TxtPVNoFr) Then
        strAppend = strAppend & " and PVNo >= '" & Trim(TxtPVNoFr) & "'"
    End If
        
    If Len(TxtPVNoTo) Then
        strAppend = strAppend & " and PVNo <= '" & Trim(TxtPVNoTo) & "'"
    End If
                
        sql = ""
        sql = "SELECT CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISDB.dbo.CISCas.DateVisit, CISDB.dbo.CISCas.Remarks, " & _
                "CISDB.dbo.CISCas.Pending, CISDB.dbo.CISCas.Claimability, CISDB.dbo.CISCas.CASStatus, dbo.Payment2MBM.PVNo, dbo.Payment2MBM.PVMBMIndex, " & _
                "dbo.Payment2MBM.PVDate, dbo.Payment2MBM.PVChequeNo, dbo.Payment2MBM.PVChequeDate, dbo.Payment2MBM.PAYAccPeriodMth, " & _
                "dbo.Payment2MBM.MBMPaidAmt, dbo.Payment2MBM.PAYAccPeriodYr, dbo.Payment2MBM.PVPayMBMStatus, dbo.Payment2MBM.Pay2MBMLastUpdateDate, " & _
                "dbo.Payment2MBM.Pay2MBMLastUpdateUser, HISDB.dbo.CLN.CLNName, CISDB.dbo.CISCas.ClinicCode, CISWorksheet.dbo.Claim.DebitCreditRefNo, " & _
                "CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVEnteredBy, CISWorksheet.dbo.ClaimInvoice.BillNo, " & _
                "CISWorksheet.dbo.Claim.CASNumber, CISWorksheet.dbo.Claim.BordxRefNo, CISWorksheet.dbo.Claim.CLMPayableByMedix2M, " & _
                "CISWorksheet.dbo.Claim.ClaimCloseDate, dbo.Payment2MBM.PVPayto, CISWorksheet.dbo.Bordx.GroupCompany, dbo.Payment2MBM.PVPayeeBank, " & _
                "dbo.Payment2MBM.PVPayeeACNo, dbo.Payment2MBM.Remark, dbo.Payment2MBM.PVCLNChequeAmt, CISDB.dbo.CISCas.OPBatchNo, " & _
                "CISWorksheet.dbo.Claim.CLMPostageFee, CISWorksheet.dbo.Claim.CLMNonGPFee, case When MBMCName IS NULL Then MBMCrossReference.MBMName else MBMCName END AS MBMName, " & _
                              "CISWorksheet.dbo.Claim.ClaimBordxStatus, dbo.Payment2MBM.PVWordingOne, dbo.Payment2MBM.PVWordingTwo, CISWorksheet.dbo.Claim.CLMFinalApprove, " & _
                              "HISDB.dbo.MBMOthersTwo.MBMBankACNo, HISDB.dbo.MBMOthersTwo.MBMBankName, CISWorksheet.dbo.Claim.CLMTotalActual, " & _
                              "HISDB.dbo.MBMOthers.MBMGroupCompany, CISDB.dbo.CISCas.InvoiceNo, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.MBMOthers.MBMCostCenter, " & _
                              "HISDB.dbo.MBMOthers.MBMSubDivision , HISDB.dbo.MBMOthers.MBMDivision, HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMEmployeeNo " & _
                "FROM dbo.Payment2MBM RIGHT OUTER JOIN CISDB.dbo.CISCas LEFT OUTER JOIN HISDB.dbo.MBMOthersTwo RIGHT OUTER JOIN " & _
                              "HISDB.dbo.MBMOthers INNER JOIN " & _
                              "HISMaintenance.dbo.MBMCrossReference ON HISDB.dbo.MBMOthers.MBMNumber = HISMaintenance.dbo.MBMCrossReference.MBMNumber ON " & _
                              "HISDB.dbo.MBMOthersTwo.MBMNumber = HISMaintenance.dbo.MBMCrossReference.MBMNumber ON " & _
                              "HISMaintenance.dbo.MBMCrossReference.MBMNumber = CISDB.dbo.CISCas.MBMnumber left outer join HISDB..MBMCoveredPersons ON MBMCoveredPersons.MBMCNumber=ciscas.MBMnumber and MBMCoveredPersons.MBMCCoverID=ciscas.CoverID RIGHT OUTER JOIN " & _
                              "CISWorksheet.dbo.Claim LEFT OUTER JOIN CISWorksheet.dbo.Bordx ON CISWorksheet.dbo.Claim.BordxRefNo = CISWorksheet.dbo.Bordx.BordxRefNo LEFT OUTER JOIN " & _
                              "CISWorksheet.dbo.ClaimInvoice ON CISWorksheet.dbo.Claim.CASNumber = CISWorksheet.dbo.ClaimInvoice.CASNumber ON " & _
                              "CISWorksheet.dbo.Claim.CASNumber = CISDB.dbo.CISCas.CASNumber ON CISWorksheet.dbo.Claim.CASNumber = dbo.Payment2MBM.CASNumber RIGHT OUTER JOIN " & _
                              "HISDB.dbo.CLN ON CISDB.dbo.CISCas.ClinicCode = HISDB.dbo.CLN.CLNCode " & _
                              "WHERE     (CISWorksheet.dbo.ClaimInvoice.INVPayTo = 'M') AND (CISDB.dbo.CISCas.CASStatus <> 'D') AND " & _
                              "(CISWorksheet.dbo.Claim.CLMPayableByMedix2M + CISWorksheet.dbo.Claim.CLMNonGPFee <> 0) And dbo.Payment2MBM.PVNo >= '" & Trim(TxtPVNoFr) & "' And dbo.Payment2MBM.PVNo <= '" & Trim(TxtPVNoTo) & "'"
                
     
    PVRec.Open sql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        
    Do While Not PVRec.EOF
        RecordFound = True
        Set Record = lsPVMBMListing.ListItems.Add(, , PVRec!PVNo)
        
        If Len(PVRec!PVDate) Then
            Record.SubItems(1) = (PVRec!PVDate)
        End If
        If Len(PVRec!MBMName) Then
            Record.SubItems(2) = (PVRec!MBMName)
        End If
        'If Len(PVRec!InvoiceNo) Then
        '    Record.SubItems(3) = (PVRec!InvoiceNo)
        'End If
        'If Len(PVRec!InvoiceDate) Then
        '    Record.SubItems(4) = (PVRec!InvoiceDate)
        'End If
        If Len(PVRec!PVChequeNo) Then
            Record.SubItems(5) = (PVRec!PVChequeNo)
        End If
        'If Len(PVRec!PVCLNDate) Then
        '    Record.SubItems(6) = (PVRec!PVCLNDate)
        'End If
        If Len(PVRec!PVCLNChequeAmt) Then
            Record.SubItems(7) = (PVRec!PVCLNChequeAmt)
        End If
        If Len(PVRec!CLMFinalApprove) Then
            Record.SubItems(8) = (PVRec!CLMFinalApprove)
        End If
        If Len(PVRec!CasNumber) Then
            Record.SubItems(9) = (PVRec!CasNumber)
        End If
        PVRec.MoveNext
    Loop
    PVRec.Close
    Set PVRec = Nothing
    
    CmdExit.Enabled = True
    
    If RecordFound = True Then
        'If CDbl(lblTotalChqList) = CDbl(lblTotalchqSel) Then
            CmdPrint.Enabled = True
            cmdprintPDF.Enabled = True
        'Else
        '    cmdPrint.Enabled = False
        '    MsgBox "Cheque range not in series", vbOKOnly + vbCritical
        'End If
    Else
        MsgBox "No record found", vbOKOnly + vbCritical
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
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

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPrintPVRange")
    If chqMode = "MedixPayMBM" Then
        lsPVListing.Visible = False
        lsPVMBMListing.Visible = True
    ElseIf chqMode = "MedixPayCLN" Then
        lsPVListing.Visible = True
        lsPVMBMListing.Visible = False
    End If
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
Private Sub txtPVNoFr_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub txtPVNoTo_Click()
    CmdPrint.Enabled = False
End Sub
Private Sub PVPrintingCLN_New()
Const METHOD_NAME = "PVPrintingCLN_New"
On Error GoTo debugs

'On Error Resume Next
Dim X As Boolean
Dim DocSavePath As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim i As Integer
Dim StrSql As String
Dim strsqlcnt As String
Dim RecordFound As Boolean
Dim PVRec As New ADODB.Recordset
Dim PVRec1 As New ADODB.Recordset
Dim RecFound As Boolean
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Cnt As Integer
Dim tmpPVCNTFr As String
Dim tmpPVCNTTo As String
Dim TmpPVCode As String
Dim tmpPVCnt As Integer
Dim tmpCOMPV As String


tmpNewRow = 0
tmpTotalBillAmt = 0
tmpTotalFFSAmt = 0
tmpTotalAmt = 0
tmpTotalCopayment = 0
tmpTotalClaimAmt = 0
tmpTotalClaimAmtSub = 0
Dim strsqlGrp As String
    strAppend = ""
    tmpPVCNTFr = Mid(TxtPVNoFr, 3, 9)
    tmpPVCNTTo = Mid(TxtPVNoTo, 3, 9)
    TmpPVCode = Mid(TxtPVNoFr, 1, 2)
    
    tmpPVCnt = 1
    RecordFound = False
    Do While tmpPVCNTFr <= tmpPVCNTTo
        tmpCOMPV = TmpPVCode & tmpPVCNTFr
        tmpTotalClaimAmtSub = 0
        tmpTotalBillAmt = 0
        tmpTotalFFSAmt = 0
        tmpTotalAmt = 0
        tmpTotalCopayment = 0
        tmpTotalClaimAmt = 0
        StrSql = ""
        StrSql = " SELECT CISDB.dbo.CISCas.ClinicCode, dbo.Payment2CLN.PVPayCLNStatus,dbo.Payment2CLN.CLNAccPeriodMth, dbo.Payment2CLN.CLNAccPeriodYr, HISDB.dbo.CLN.CLNName, " & _
                " CISWorksheet.dbo.ClaimInvoice.INVPayTo,HISDB.dbo.CLN.CLNAdd1, HISDB.dbo.CLN.CLNAdd2, HISDB.dbo.CLN.CLNAdd3, " & _
                " HISDB.dbo.CLN.CLNState, HISDB.dbo.CLN.CLNCity, HISDB.dbo.CLN.CLNPostCode, dbo.Payment2CLN.PVCLNNo,dbo.Payment2CLN.PVCLNDate, dbo.Payment2CLN.PVCLNChequeDate, dbo.Payment2CLN.PVCLNChequeNo, CISDB.dbo.CISCas.InvoiceNo, " & _
                " CISDB.dbo.CISCas.InvoiceDate, dbo.Payment2CLN.PVCLNChequeAmt, CISDB.dbo.CISCas.CASStatus, dbo.Payment2CLN.PVCLNBank,dbo.Payment2CLN.PVCLNAccount,  dbo.Payment2CLN.CLNRemark, HISDB.dbo.CLN.ClinicGrpCode, " & _
                " HISDB.dbo.CLN.CLNPName, HISDB.dbo.CLN.CLNPAdd1,HISDB.dbo.CLN.CLNPAdd2, HISDB.dbo.CLN.CLNPAdd3, HISDB.dbo.CLN.CLNEmail, HISDB.dbo.CLN.CLNBankAcc, dbo.Payment2CLN.PVWordingOne,dbo.Payment2CLN.PVWordingTwo," & _
                " HISDB.dbo.CLN.CLNType, HISDB.dbo.CLN.CLNFFSStatus, HISDB.dbo.CLN.CLNPayee, HISDB.dbo.CLN.CLNRegNo, HISDB.dbo.CLN.CLNBankName, SUM(CISWorksheet.dbo.Claim.CLMTotalActual) AS CLMTotalActual,sum(CISWorksheet.dbo.Claim.CLMCoPayCLNColl) as Copay, " & _
                " sum(CISWorksheet.dbo.Claim.CLMTotalApproved ) as TotalApproved,SUM(CISWorksheet.dbo.Claim.CLMMCOProcFee) AS CLMMCOProcFee, SUM(CISWorksheet.dbo.Claim.CLMBordxAmt ) AS CLMBordxAmt,SUM(HISDB.dbo.CLN.CLNMCOProcfeePerc) AS CLNMCOProcfeePerc," & _
                " SUM(CISWorksheet.dbo.Claim.CLMPayableByMedix2C) AS CLMPayableByMedix2C,SUM(dbo.Payment2CLN.CLNPaidAmt) AS CLNPaidAmt,SUM(CISWorksheet.dbo.Claim.CLMNonGPFee) AS CLMNonGPFee ,SUM(CISWorksheet.dbo.Claim.CLMFinalApprove) AS CLMFinalApprove" & _
                " FROM HISDB.dbo.CLN INNER JOIN " & _
                " CISDB.dbo.CISCas ON HISDB.dbo.CLN.CLNCode = CISDB.dbo.CISCas.ClinicCode RIGHT OUTER JOIN   " & _
                " dbo.Payment2CLN RIGHT OUTER JOIN   " & _
                " CISWorksheet.dbo.Bordx RIGHT OUTER JOIN   " & _
                " CISWorksheet.dbo.ClaimInvoice INNER JOIN   " & _
                " CISWorksheet.dbo.Claim ON CISWorksheet.dbo.ClaimInvoice.CASNumber = CISWorksheet.dbo.Claim.CASNumber ON   " & _
                " CISWorksheet.dbo.Bordx.BordxRefNo = CISWorksheet.dbo.Claim.BordxRefNo ON dbo.Payment2CLN.CASNumber = CISWorksheet.dbo.Claim.CASNumber ON   " & _
                " CISDB.dbo.CISCAS.CasNumber = CISWorksheet.dbo.Claim.CasNumber  where CISDB.dbo.CISCas.InvoiceNo <>N'' and  CISWorksheet.dbo.ClaimInvoice.INVPayTo = 'C' And CLMFinalApprove is not null AND dbo.Payment2CLN.PVCLNNo = '" & Trim(tmpCOMPV) & "' And Claimability <> 'C'"
                    
                strsqlGrp = "GROUP BY CISDB.dbo.CISCas.ClinicCode, dbo.Payment2CLN.PVPayCLNStatus, dbo.Payment2CLN.CLNAccPeriodMth, dbo.Payment2CLN.CLNAccPeriodYr, " & _
                " HISDB.dbo.CLN.CLNName, CISWorksheet.dbo.ClaimInvoice.INVPayTo,  HISDB.dbo.CLN.CLNAdd1,   " & _
                " HISDB.dbo.CLN.CLNAdd2, HISDB.dbo.CLN.CLNAdd3, HISDB.dbo.CLN.CLNState, HISDB.dbo.CLN.CLNCity, HISDB.dbo.CLN.CLNPostCode, dbo.Payment2CLN.PVCLNNo, " & _
                " dbo.Payment2CLN.PVCLNDate, dbo.Payment2CLN.PVCLNChequeDate, dbo.Payment2CLN.PVCLNChequeNo, CISDB.dbo.CISCas.InvoiceNo, " & _
                " CISDB.dbo.CISCas.InvoiceDate, dbo.Payment2CLN.PVCLNChequeAmt, CISDB.dbo.CISCas.CASStatus, dbo.Payment2CLN.PVCLNBank, " & _
                " dbo.Payment2CLN.PVCLNAccount, dbo.Payment2CLN.CLNRemark, HISDB.dbo.CLN.ClinicGrpCode, HISDB.dbo.CLN.CLNPName, HISDB.dbo.CLN.CLNPAdd1, " & _
                " HISDB.dbo.CLN.CLNPAdd2, HISDB.dbo.CLN.CLNPAdd3, HISDB.dbo.CLN.CLNEmail, HISDB.dbo.CLN.CLNBankAcc, dbo.Payment2CLN.PVWordingOne, " & _
                " dbo.Payment2CLN.PVWordingTwo, HISDB.dbo.CLN.CLNType, " & _
                " HISDB.dbo.CLN.CLNFFSStatus , HISDB.dbo.CLN.CLNPayee, HISDB.dbo.CLN.CLNRegNo, HISDB.dbo.CLN.CLNBankName "
                StrSql = StrSql & strsqlGrp
        PVRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        
        strsqlcnt = ""
        strsqlcnt = " SELECT Count(distinct CISCas.InvoiceNo) AS CNT_Rec FROM HISDB.dbo.CLN INNER JOIN CISDB.dbo.CISCas ON HISDB.dbo.CLN.CLNCode = CISDB.dbo.CISCas.ClinicCode RIGHT OUTER JOIN   dbo.Payment2CLN RIGHT OUTER JOIN   CISWorksheet.dbo.Bordx RIGHT OUTER JOIN   CISWorksheet.dbo.ClaimInvoice INNER JOIN    " & _
                    " CISWorksheet.dbo.Claim ON CISWorksheet.dbo.ClaimInvoice.CASNumber = CISWorksheet.dbo.Claim.CASNumber ON   CISWorksheet.dbo.Bordx.BordxRefNo = CISWorksheet.dbo.Claim.BordxRefNo ON dbo.Payment2CLN.CASNumber = CISWorksheet.dbo.Claim.CASNumber ON   " & _
                    " CISDB.dbo.CISCAS.CasNumber = CISWorksheet.dbo.Claim.CasNumber  where CISDB.dbo.CISCas.InvoiceNo <>N'' and  CISWorksheet.dbo.ClaimInvoice.INVPayTo = 'C' And CLMFinalApprove is not null AND dbo.Payment2CLN.PVCLNNo = '" & Trim(tmpCOMPV) & "' And Claimability <> 'C'"
        PVRec1.Open strsqlcnt, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
        
        tmp_Rec = 0
        tmp_Rec = PVRec1!CNT_Rec
        
        PVRec1.Close
        Set PVRec1 = Nothing
        
        RecordFound = False
        Cnt = 0
        tmpTotalFFSAmt = 0
        
        Dim checkno As String
        Dim ForCount As Integer
        ForCount = 1
        Do While Not PVRec.EOF
            RecordFound = True
            If objExcel Is Nothing Then
                Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
                objExcel.Visible = False
                objExcel.WindowState = xlMaximized
                
                If Mid(PVRec!PVCLNNo, 1, 2) <> "CS" Then
                    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-ClinicTest.xlt")
                ElseIf Mid(PVRec!PVCLNNo, 1, 2) = "CS" Then
                    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\PV-Clinic_SihatJohor.xlt")
                End If
                
                Set objWorksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objWorksheet = objWorkbook.Worksheets.Item(1)
            End If
                Cnt = CDbl(Cnt) + 1
                If Cnt = 1 Then
                    objWorksheet.Cells(8, "C") = IIf(Len(PVRec!CLNPName), PVRec!CLNPName, "")
                    objWorksheet.Cells(9, "C") = PVRec!CLNName & " (" & PVRec!ClinicCode & ")"
                    objWorksheet.Cells(10, "C") = IIf(Len(PVRec!CLNPAdd1), PVRec!CLNPAdd1, "")
                    objWorksheet.Cells(11, "C") = IIf(Len(PVRec!CLNPAdd2), PVRec!CLNPAdd2, "")
                    objWorksheet.Cells(12, "C") = IIf(Len(PVRec!CLNPAdd3), PVRec!CLNPAdd3, "")
                    objWorksheet.Cells(8, "H") = PVRec!PVCLNNo '& "/" & Format(PVRec!PVCLNDate, "MM/YY")
                    objWorksheet.Cells(9, "H") = Trim(PVRec!PVCLNChequeNo)
                    objWorksheet.Cells(10, "H") = PVRec!PVCLNDate
                    tmpNewRow = 17 + CDbl(Cnt)
                    If Trim(TxtPVNoFr) = Trim(TxtPVNoTo) And Len(PVRec!CLNRemark) <> 0 Then
                        objWorksheet.Range("A14" + ":" + "J14").RowHeight = 30
                        objWorksheet.Cells(14, "A") = "Remark : " + Trim(PVRec!CLNRemark)
                    End If
                End If
                
                If CDbl(tmpNewRow) <= "44" Then
                    objWorksheet.Cells(16 + CDbl(Cnt), "A") = Cnt
                    objWorksheet.Cells(16 + CDbl(Cnt), "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                    objWorksheet.Cells(16 + CDbl(Cnt), "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "D") = PVRec!InvoiceDate
                    objWorksheet.Cells(16 + CDbl(Cnt), "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "F") = PVRec!CLMTotalActual
                    objWorksheet.Cells(16 + CDbl(Cnt), "F").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    objWorksheet.Cells(16 + CDbl(Cnt), "G").Borders.LineStyle = XlLineStyle.xlContinuous
                    'objWorksheet.Cells(16 + CDbl(Cnt), "H") = IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                    objWorksheet.Cells(16 + CDbl(Cnt), "H") = CDbl(PVRec!CLMTotalActual) - CDbl(PVRec!copay)
                    tmpTotalClaimAmtSub = 0
                    tmpTotalClaimAmtSub = CDbl(PVRec!CLMTotalActual) - CDbl(PVRec!copay)
                    objWorksheet.Cells(16 + CDbl(Cnt), "H").Borders.LineStyle = XlLineStyle.xlContinuous
                    
                    objWorksheet.Cells(16 + CDbl(Cnt), "I") = IIf(Len(PVRec!CLMMCOProcFee), PVRec!CLMMCOProcFee, 0)
                    objWorksheet.Cells(16 + CDbl(Cnt), "I").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(16 + CDbl(Cnt), "J") = IIf(Len(PVRec!CLMPayableByMedix2C), CDbl(PVRec!CLMPayableByMedix2C), 0)
                    objWorksheet.Cells(16 + CDbl(Cnt), "J").Borders.LineStyle = XlLineStyle.xlContinuous
                                        
                    Dim Fnlapp As Double
                    If IsNull(PVRec!CLMTotalActual) Then
                        Fnlapp = 0
                    Else
                        Fnlapp = PVRec!CLMTotalActual
                    End If
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + Fnlapp
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = CDbl(tmpTotalClaimAmt) + CDbl(tmpTotalClaimAmtSub)
                    tmpNewRow = CDbl(tmpNewRow) + 1
                    
                 ElseIf CDbl(tmpNewRow) > "44" Then
                   
                    Do While Cnt - 1 < tmp_Rec
                    ForCount = ForCount + 1
                    
                        Call Bind_Page_Excelvlaues(objWorksheet, PVRec, tmpNewRow, Cnt, ForCount)
                    Loop
                    If PVRec.EOF Then
                        Exit Do
                    End If
            End If
            PVRec.MoveNext
        Loop
                If (Cnt < 27) Then
                    tmpNewRow = tmpNewRow + 0
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "F") = tmpTotalBillAmt
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "F").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "G") = tmpTotalCopayment
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "G").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "H") = tmpTotalClaimAmt
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "H").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "I") = tmpTotalFFSAmt
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "I").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "J") = tmpTotalAmt
                    objWorksheet.Cells(CDbl(tmpNewRow) - 1, "J").Borders.LineStyle = XlLineStyle.xlContinuous
                    'tmpNewRow = tmpNewRow + 5
                    objWorksheet.Cells(44, "C") = "Prepared By : " & Logon
                    objWorksheet.Cells(44, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(44 + 2, "C") = "Approved By : "
                    objWorksheet.Cells(44 + 2, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(44 + 4, "C") = "Signatories : "
                    objWorksheet.Cells(44 + 4, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(44, "G") = "Acknowledgement of receipt : "
                    objWorksheet.Cells(44 + 3, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(44 + 3, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(44 + 4, "G") = "Signature & Company Stamp"
                Else
                    tmpNewRow = tmpNewRow
                    objWorksheet.Cells(CDbl(tmpNewRow), "F") = tmpTotalBillAmt
                    objWorksheet.Cells(CDbl(tmpNewRow), "F").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "G") = tmpTotalCopayment
                    objWorksheet.Cells(CDbl(tmpNewRow), "G").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "H") = tmpTotalClaimAmt
                    objWorksheet.Cells(CDbl(tmpNewRow), "H").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "I") = tmpTotalFFSAmt
                    objWorksheet.Cells(CDbl(tmpNewRow), "I").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "J") = tmpTotalAmt
                    objWorksheet.Cells(CDbl(tmpNewRow), "J").Borders.LineStyle = XlLineStyle.xlContinuous
                    
                    If ForCount >= 2 Then
                    Dim NrCnt As Integer
                    NrCnt = ForCount * 51
                    tmpNewRow = NrCnt - 6
                    objWorksheet.Cells(tmpNewRow, "C") = "Prepared By : " & Logon
                    objWorksheet.Cells(tmpNewRow, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(tmpNewRow + 2, "C") = "Approved By : "
                    objWorksheet.Cells(tmpNewRow + 2, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(tmpNewRow + 4, "C") = "Signatories : "
                    objWorksheet.Cells(tmpNewRow + 4, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(tmpNewRow, "G") = "Acknowledgement of receipt : "
                    objWorksheet.Cells(tmpNewRow + 3, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(tmpNewRow + 3, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(tmpNewRow + 4, "G") = "Signature & Company Stamp"
                    End If
                    'tmpNewRow = tmpNewRow + 6
                    'objWorksheet.Cells(tmpNewRow, "C") = "Prepared By : " & Logon
                    'objWorksheet.Cells(tmpNewRow, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    'objWorksheet.Cells(tmpNewRow + 2, "C") = "Approved By : "
                    'objWorksheet.Cells(tmpNewRow + 2, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    'objWorksheet.Cells(tmpNewRow + 4, "C") = "Signatories : "
                    'objWorksheet.Cells(tmpNewRow + 4, "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    'objWorksheet.Cells(tmpNewRow, "G") = "Acknowledgement of receipt : "
                    'objWorksheet.Cells(tmpNewRow + 3, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    'objWorksheet.Cells(tmpNewRow + 3, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    'objWorksheet.Cells(tmpNewRow + 4, "G") = "Signature & Company Stamp"
                End If
                
        PVRec.Close
        Set PVRec = Nothing
        
        
        
        If RecordFound = True Then
            Dim fPDF As String
            
            If PrintPDFstatus = True Then
            'Template
                X = CreatePath(DocPath & "ClinicPVS\" & Trim(tmpCOMPV))
                DocSavePath = DocPath & "ClinicPVS\" & Trim(tmpCOMPV) & "\"
                fPDF = DocPath & "ClinicPVS\" & Trim(tmpCOMPV) & "\" & Trim(tmpCOMPV) & ".pdf"
                objWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fPDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
                'objWorkbook.SaveAs fPDF
            Else
                objWorkbook.PrintOut
                fPDF = "C:\Template\LatestPV.xls"
                objWorkbook.SaveAs fPDF
            End If
            
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            Set objExcel = Nothing
        Else
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            Set objExcel = Nothing

        End If
        tmpPVCNTFr = Format(CDbl(tmpPVCNTFr) + 1, "000000000")
       
    Loop
ExitFunction:
            txtPrintingStatus = "Printing Not In Progress!"
            CmdPrint.Enabled = False
            CmdExit.Enabled = True
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            txtPrintingStatus = "Printing Not In Progress!"
            CmdPrint.Enabled = True
            CmdExit.Enabled = True
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
debugs:
If Err.Description <> "" Then
   Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub Bind_Page_Excelvlaues(objWorksheet As Excel.Worksheet, PVRec As ADODB.Recordset, tmpNewRow As Double, temprec As Integer, ForCount As Integer)
Dim IsFrst As Boolean
Dim Rcnt As Integer
Dim Cnt As Integer
Cnt = 0
Rcnt = 0
IsFrst = True
    For Rcnt = 0 To 37
     
        
            If PVRec.EOF Then
                Exit Sub
            Else
                    If IsFrst Then
                        'tmpNewRow = CDbl(tmpNewRow) + 13
                        tmpNewRow = 51 * (ForCount - 1) + 5
                        objWorksheet.Cells(CDbl(tmpNewRow), "A") = "No"
                        objWorksheet.Cells(CDbl(tmpNewRow), "A").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "C").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "C") = "Invoice No."
                        objWorksheet.Cells(CDbl(tmpNewRow), "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "D") = "Invoice Date"
                        objWorksheet.Cells(CDbl(tmpNewRow), "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "F") = "Amount"
                        objWorksheet.Cells(CDbl(tmpNewRow), "F").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "G") = "FFS"
                        objWorksheet.Cells(CDbl(tmpNewRow), "G").Borders.LineStyle = XlLineStyle.xlContinuous
                        objWorksheet.Cells(CDbl(tmpNewRow), "H").Borders.LineStyle = XlLineStyle.xlContinuous
                        tmpNewRow = CDbl(tmpNewRow) + 1
                        IsFrst = False
                    End If
                    
                   
                    objWorksheet.Cells(CDbl(tmpNewRow), "A") = temprec
                    objWorksheet.Cells(CDbl(tmpNewRow), "A").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "A").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "B").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "B").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "C") = IIf(Len(PVRec!InvoiceNo), Trim(PVRec!InvoiceNo), "")
                    objWorksheet.Cells(CDbl(tmpNewRow), "C").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    
                    objWorksheet.Cells(CDbl(tmpNewRow), "D") = PVRec!InvoiceDate
                    objWorksheet.Cells(CDbl(tmpNewRow), "D").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "D").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    
                    objWorksheet.Cells(CDbl(tmpNewRow), "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "E").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "E").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "F") = PVRec!CLMTotalActual
                    objWorksheet.Cells(CDbl(tmpNewRow), "F").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "G") = IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    objWorksheet.Cells(CDbl(tmpNewRow), "G").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "H") = PVRec!CLMBordxAmt
                    objWorksheet.Cells(CDbl(tmpNewRow), "H").Borders.LineStyle = XlLineStyle.xlContinuous
                    
                    objWorksheet.Cells(CDbl(tmpNewRow), "I") = PVRec!CLMMCOProcFee
                    objWorksheet.Cells(CDbl(tmpNewRow), "I").Borders.LineStyle = XlLineStyle.xlContinuous
                    objWorksheet.Cells(CDbl(tmpNewRow), "J") = CDbl(PVRec!CLMPayableByMedix2C)
                    objWorksheet.Cells(CDbl(tmpNewRow), "J").Borders.LineStyle = XlLineStyle.xlContinuous
                    
                    tmpTotalBillAmt = IIf(Len(tmpTotalBillAmt), CDbl(tmpTotalBillAmt), 0) + IIf(Len(PVRec!CLMTotalActual), PVRec!CLMTotalActual, 0)
                    tmpTotalFFSAmt = IIf(Len((tmpTotalFFSAmt)), CDbl(tmpTotalFFSAmt), 0) + IIf(Len((PVRec!CLMMCOProcFee)), CDbl(PVRec!CLMMCOProcFee), 0)
                    tmpTotalAmt = IIf(Len((tmpTotalAmt)), CDbl(tmpTotalAmt), 0) + IIf(Len(PVRec!CLNPaidAmt), CDbl(PVRec!CLNPaidAmt), 0)
                    tmpTotalCopayment = tmpTotalCopayment + IIf(IsNumeric(PVRec!copay), PVRec!copay, 0)
                    tmpTotalClaimAmt = tmpTotalClaimAmt + IIf(Len(PVRec!CLMBordxAmt), CDbl(PVRec!CLMBordxAmt), 0)
                    
                        
                    tmpNewRow = CDbl(tmpNewRow) + 1
                    PVRec.MoveNext
            End If
     temprec = temprec + 1
    Next
End Sub





