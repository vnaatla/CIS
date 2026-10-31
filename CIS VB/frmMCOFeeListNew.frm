VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCOFeeListNew 
   Caption         =   "MCO Fee Listing"
   ClientHeight    =   7812
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11856
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7812
   ScaleWidth      =   11856
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameMCO 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   11655
      Begin Crystal.CrystalReport cRptMCO 
         Left            =   10440
         Top             =   240
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         UserName        =   "sa"
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_MCO 
         Left            =   9960
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboMCOListingVer 
         Height          =   315
         ItemData        =   "frmMCOFeeListNew.frx":0000
         Left            =   2160
         List            =   "frmMCOFeeListNew.frx":000D
         TabIndex        =   0
         Top             =   285
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOGrpCom 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOPayor 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   4215
      End
      Begin VB.ComboBox CboPrintStatus 
         Height          =   315
         ItemData        =   "frmMCOFeeListNew.frx":0066
         Left            =   7680
         List            =   "frmMCOFeeListNew.frx":0070
         TabIndex        =   9
         Top             =   285
         Width           =   1335
      End
      Begin VB.ComboBox CboRptType 
         Height          =   315
         ItemData        =   "frmMCOFeeListNew.frx":0085
         Left            =   7680
         List            =   "frmMCOFeeListNew.frx":008F
         TabIndex        =   10
         Top             =   555
         Width           =   1335
      End
      Begin VB.ComboBox cboMCOCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCOFeeListNew.frx":00AD
         Left            =   2160
         List            =   "frmMCOFeeListNew.frx":00BA
         TabIndex        =   3
         Top             =   840
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtMCOBordxDateFr 
         Height          =   315
         Left            =   7680
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   315
         Left            =   7680
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtMCOBordxDate 
         Height          =   315
         Left            =   7680
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMCOBordxDateTo 
         Height          =   315
         Left            =   10080
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtMCOBatchNo 
         Height          =   315
         Left            =   10080
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Print Status"
         Height          =   255
         Left            =   6600
         TabIndex        =   30
         Top             =   320
         Width           =   975
      End
      Begin VB.Label lblDisplay1 
         Height          =   255
         Left            =   6600
         TabIndex        =   29
         Top             =   880
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Report Type"
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblDisplay2 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   9000
         TabIndex        =   27
         Top             =   880
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Single/Batch/Range"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label lbl20 
         Caption         =   "Group Company"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lbl22 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Fee Listing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox lblSupp 
         Height          =   285
         Left            =   3360
         TabIndex        =   32
         Text            =   "0"
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox lblPrinc 
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Text            =   "0"
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdMCOSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   300
         Left            =   6720
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   7560
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrintPreview 
         Caption         =   "Pre&view"
         Height          =   300
         Left            =   8400
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9240
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label201 
         Alignment       =   2  'Center
         Caption         =   "Total Supp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2400
         TabIndex        =   20
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label200 
         Alignment       =   2  'Center
         Caption         =   "Total Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5040
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   11655
      Begin MSComctlLib.ListView lvwMCO 
         Height          =   4650
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11415
         _ExtentX        =   20130
         _ExtentY        =   8192
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "CovID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "BordxDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Eff Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "ExpDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Health Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Batch"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   " "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Department"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Printed"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "InvoiceNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "InvoiceDate"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMCOFeeListNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim MultiInvFound As Boolean

Private Sub ComboKeyPress(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(tmpComboBox.Text, tmpComboBox.SelStart) + Chr(KeyAscii) + Mid(tmpComboBox.Text, tmpComboBox.SelStart + tmpComboBox.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To tmpComboBox.ListCount - 1
        If NewText = LCase(Left(tmpComboBox.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = tmpComboBox.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           tmpComboBox.Text = ValidValue
           tmpComboBox.SelStart = 0
           tmpComboBox.SelLength = Len(ValidValue)
         End If
    End If
End Sub

Private Sub cboMCOListingVer_Click()
    If cboMCOListingVer.ListIndex = 0 Then
        lbl20.Caption = "Payor"
        cboMCOPayor.Visible = True
        cboMCOGrpCom.Visible = False
    Else
        lbl20.Caption = "Group Company"
        cboMCOGrpCom.Visible = True
        cboMCOPayor.Visible = False
    End If
End Sub


Private Sub CboPrintStatus_Click()
    lvwMCO.ListItems.Clear
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

Private Sub cboMCOCriteriaSel_Click()

    cboMCOPayor = ""
    Call cmdMCOClear_Click
    txtInvoiceNo.Visible = False
    txtMCOBordxDate.Visible = False
    txtMCOBatchNo.Visible = False
    txtMCOBordxDateFr.Visible = False
    txtMCOBordxDateTo.Visible = False
    txtMCOBatchNo = ""
    Select Case cboMCOCriteriaSel.ListIndex
        Case 0 'Batch
            lblDisplay1.Caption = "Bordx Date"
            lblDisplay2.Caption = "Batch No."
            txtMCOBordxDate.Visible = True
            txtMCOBatchNo.Visible = True
        Case 1 'Range
            lblDisplay1.Caption = "From"
            lblDisplay2.Caption = "To"
            txtMCOBordxDateFr.Visible = True
            txtMCOBordxDateTo.Visible = True
        Case 2 'Invoice No
            lblDisplay1.Caption = "Invoice No."
            lblDisplay2.Caption = ""
            txtInvoiceNo.Visible = True
    End Select
End Sub

Private Sub cboMCOListingVer_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMCOListingVer, KeyAscii)
End Sub

Private Sub cboMCOPayor_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMCOPayor, KeyAscii)
End Sub

Private Sub cmdMCOClear_Click()
    txtMCOBatchNo = ""
    cboMCOPayor = ""
    txtMCOBordxDate = "__/__/____"
    txtMCOBordxDateFr = "__/__/____"
    txtMCOBordxDateTo = "__/__/____"
    'cboMCOListingVer.ListIndex = 1
    lvwMCO.ListItems.Clear
    cmdMCOPrint.Enabled = False
    cmdMCOPrintPreview.Enabled = False
    lblSupp = 0
    lblPrinc = 0
End Sub

Private Sub cmdMCOClose_Click()
    Unload Me
End Sub

Private Sub cmdMCOPrint_Click()
Dim tempPath As String
    tempPath = "\\PLUTO\HISSaturn\Conn\MediConnCISDBPluto.UDL"
    If MultiInvFound = True Then
        MsgBox "Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! ", vbCritical + vbOKOnly
        Exit Sub
    ElseIf Trim(CboRptType) = "" Then
        MsgBox "Please select Report Type ! ", vbCritical + vbOKOnly
    Else
        Screen.MousePointer = vbHourglass
        MediConnCISDBPluto.Open "File Name=" & tempPath
        Call CrystalRptName
        MediConnCISDBPluto.Close
        Set MediConnCISDBPluto = Nothing
        Screen.MousePointer = vbDefault
        cRptMCO.Reset
    End If
End Sub

Private Sub cmdMCOPrintPreview_Click()
Dim tempPath As String
    tempPath = "\\PLUTO\HISSaturn\Conn\MediConnCISDBPluto.UDL"
    If MultiInvFound = True Then
        MsgBox "Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! ", vbCritical + vbOKOnly
        Exit Sub
    ElseIf Trim(CboRptType) = "" Then
        MsgBox "Please select Report Type ! ", vbCritical + vbOKOnly
    Else
        Screen.MousePointer = vbHourglass
        MediConnCISDBPluto.Open "File Name=" & tempPath
        Call CrystalRptName
        MediConnCISDBPluto.Close
        Set MediConnCISDBPluto = Nothing
        Screen.MousePointer = vbDefault
        cRptMCO.Reset
    End If
End Sub

Private Sub cmdMCOSearch_Click()
Const METHOD_NAME = "cmdMCOSearch_Click"
On Error GoTo debugs

    lblPrinc = ""
    lblSupp = ""

    If cboMCOCriteriaSel = "" Then
        MsgBox "Please Enter MCO Criteria!", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    If cboMCOCriteriaSel.ListIndex = 2 Then
        If Trim(txtInvoiceNo) = "" Then
            MsgBox "Please Enter Invoice Number!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    Else
        If cboMCOCriteriaSel.ListIndex = 0 Then
            If txtMCOBordxDate = "__/__/____" Or txtMCOBatchNo = "" Then
                MsgBox "Please Enter Bordereaux Date and Batch No!", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            End If
        ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
            If txtMCOBordxDateFr = "__/__/____" Or txtMCOBordxDateTo = "__/__/____" Then
                MsgBox "Please Enter Bordereaux Range!", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            End If
        End If
        If cboMCOListingVer = "" Then
            MsgBox "Please Enter Listing Version!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
        If cboMCOListingVer.ListIndex = 0 Then
            If cboMCOPayor = "" Then
                MsgBox "Please Enter Payor!", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            End If
        Else
            If cboMCOGrpCom = "" Then
                MsgBox "Please Enter Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            End If
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    Call MCOFeeListing
    Screen.MousePointer = Default
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub MCOFeeListing()
Const METHOD_NAME = "MCOFeeListing"
On Error GoTo debugs
Dim rsSearch As New ADODB.Recordset
Dim TotPrinc As Integer
Dim TotSupp As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim StrSql As String
Dim strCriterial As String
Dim strListingV As String
Dim tmpSleCri As String
Dim PrevInvNO As String
Dim strSelFields As String
Dim strSortlist As String
Dim ListMaster As ListItem
Dim CurrInvNo As String
    TotPrinc = 0
    TotSupp = 0
    strCriterial = ""
    strListingV = ""
    MultiInvFound = False
    PrevInvNO = ""
    lvwMCO.ListItems.Clear
   
    strSelFields = "SELECT MBMCCoverID, PAYCode, MBMPolEffDate, MBMPolExpDate, MBMNumber,  " & _
                   " MBMBordxDate, MBMCRelCode, MBMBatchNo, MBMDepartment, INSCode, GRPCoCode, " & _
                   "MBMMemberType, MBMName, MBMGroupCompany, MBMCInvoiceNo, MBMCMCOFee,MBMCInvoiceDate "
    tmpSleCri = ""
    If cboMCOCriteriaSel.ListIndex = 2 Then
        strCriterial = " AND MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "'"
    ElseIf cboMCOCriteriaSel.ListIndex = 0 Then
        strCriterial = " AND (MBMBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "'" & _
                   " AND MBMBatchNo = " & Trim(txtMCOBatchNo) & ")"
    ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
        strCriterial = " AND MBMBordxDate >= '" & Format(txtMCOBordxDateFr, "mm/dd/yyyy") & "'" & _
              " AND MBMBordxDate <= '" & Format(txtMCOBordxDateTo, "mm/dd/yyyy") & "'"
    End If
    If cboMCOCriteriaSel.ListIndex <> 2 Then
        If CboPrintStatus.ListIndex = 0 Then
            strCriterial = strCriterial & " And MBMCInvoiced='N' "
        Else
            strCriterial = strCriterial & " And MBMCInvoiced ='Y'"
        End If
        If cboMCOListingVer = "MCO Listing - by Payor" Then
            strListingV = " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
        Else
            strListingV = " AND GrpCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
        End If
    End If
    tmpSleCri = strCriterial & strListingV

    tmpSleCri = tmpSleCri & " ORDER BY MBMNumber"
    strSortlist = ""
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "MCOFeeNew"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, strSelFields)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, tmpSleCri)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("SelSortList", adVarChar, adParamInput, 2000, strSortlist)
    cmd.Parameters.Append Prm3
    Set Prm4 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 5, 4)
    cmd.Parameters.Append Prm4

    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    rsSearch.CacheSize = 100
    Set rsSearch = cmd.Execute
    cmdMCOPrint.Enabled = True
    Do Until rsSearch.EOF
        If rsSearch!MBMCRELCode = "P" Then
            TotPrinc = TotPrinc + 1
        Else
            TotSupp = TotSupp + 1
        End If
        
        Set ListMaster = lvwMCO.ListItems.Add(, , rsSearch!PAYCode)
        ListMaster.SubItems(1) = rsSearch!MBMNumber
        ListMaster.SubItems(2) = rsSearch!MBMCCoverID
        ListMaster.SubItems(3) = rsSearch!MBMCRELCode
        ListMaster.SubItems(4) = rsSearch!MBMBordxDate
        ListMaster.SubItems(5) = rsSearch!MBMPolEffDate
        ListMaster.SubItems(6) = rsSearch!MBMPolEXPDate
        ListMaster.SubItems(7) = IIf(IsNumeric(rsSearch!MBMCMCOFee), rsSearch!MBMCMCOFee, 0)
        ListMaster.SubItems(8) = IIf(Len(rsSearch!MBMBatchNo), rsSearch!MBMBatchNo, "")
        ListMaster.SubItems(9) = IIf(Len(rsSearch!MBMName), Trim(rsSearch!MBMName), "")
        ListMaster.SubItems(10) = IIf(Len(Trim(rsSearch!MBMGroupCompany)), Trim(rsSearch!MBMGroupCompany), "")
        ListMaster.SubItems(11) = IIf(Len(Trim(rsSearch!GRPCoCode)), Trim(rsSearch!GRPCoCode), "")
        ListMaster.SubItems(12) = IIf(Len(Trim(rsSearch!MBMDepartment)), Trim(rsSearch!MBMDepartment), "")
        ListMaster.SubItems(13) = "NO"
        If Trim(rsSearch!MBMCInvoiceNo) <> "" Then
            ListMaster.SubItems(13) = "YES"
        End If
        If Len(rsSearch!MBMCInvoiceNo) Then
            cmdMCOPrint.Enabled = False
            CurrInvNo = Trim(rsSearch!MBMCInvoiceNo)
            If PrevInvNO = "" Then
                PrevInvNO = Trim(rsSearch!MBMCInvoiceNo)
            ElseIf CurrInvNo <> PrevInvNO Then
                MultiInvFound = True
            End If
        End If
        ListMaster.SubItems(14) = IIf(Len(rsSearch!MBMCInvoiceNo), Trim(rsSearch!MBMCInvoiceNo), "")
        ListMaster.SubItems(15) = IIf(Len(rsSearch!MBMCInvoiceDate), rsSearch!MBMCInvoiceDate, "")
        
        lblPrinc = TotPrinc
        lblSupp = TotSupp
        cmdMCOPrintPreview.Enabled = True
        rsSearch.MoveNext
    Loop
    
    If lvwMCO.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdMCOPrint.Enabled = False
        cmdMCOPrintPreview.Enabled = False
        Screen.MousePointer = Default
    End If
    
    rsSearch.Close
    Set rsSearch = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CrystalRptName()
Const METHOD_NAME = "CrystalRptName"
On Error GoTo debugs

Dim StrFormula As String
Dim tmpBatchNo As String
Dim tmpPayCode As String
Dim tmpDept As String
Dim strsqlPlutoMCO As String
Dim PlutoMCO As New ADODB.Recordset
Dim PLUTOStatus As Boolean
Dim StrSql As String
Dim rst As New ADODB.Recordset
Dim sqlinsert As String

Dim tmpBordxDate As String
Dim tmpPolEffDate As String
Dim tmpPolExpDate As String
Dim tmpInvoiceDate As String
Dim tmpMBMName As String
Dim tmpMBMCName As String
Dim tmpMBMBasicMCO As Double
Dim tmpMBMMCO As Double
Dim tmpMBMMCOFee As Double
Dim tmpPremium As Double
Dim tmpGRPCompany As String
Dim tmpDepartment As String

    Select Case cboMCOCriteriaSel.ListIndex ' By Batch
        Case 0
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(8)
        StrFormula = ""
            StrFormula = " And MBMBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "'" & _
                    " And MBMBatchNo = '" & tmpBatchNo & "'"
        Case 1
            StrFormula = " And MBMBordxDate >= '" & Format(txtMCOBordxDateFr, "mm/dd/yyyy") & "'" & _
                " And MBMBordxDate <= '" & Format(txtMCOBordxDateTo, "mm/dd/yyyy") & "'"
        Case 2
            StrFormula = " And MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "'"
    End Select
    
    If cboMCOCriteriaSel.ListIndex <> 2 Then
        If CboPrintStatus.ListIndex = 0 Then
            StrFormula = StrFormula & " And MBMCInvoiced = 'N' "
        Else
            StrFormula = StrFormula & " And MBMCInvoiced = 'Y' "
        End If
        If cboMCOListingVer = "MCO Listing - by Payor" Then
            tmpPayCode = lvwMCO.SelectedItem
            StrFormula = StrFormula & " AND PayCode = '" & tmpPayCode & "'"
        Else
            tmpPayCode = Mid(cboMCOGrpCom, 1, 4)
            StrFormula = StrFormula & " AND  GrpCoCode = '" & tmpPayCode & "'"
        End If
    End If
   
    strsqlPlutoMCO = "Select MBMNumber from MCOInvoicePluto where 0 = 0  "
    strsqlPlutoMCO = strsqlPlutoMCO + StrFormula
    PlutoMCO.Open strsqlPlutoMCO, MediConnCISDBPluto, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PlutoMCO.EOF
        PLUTOStatus = True
        PlutoMCO.MoveNext
    Loop
    PlutoMCO.Close
    Set PlutoMCO = Nothing
    
    If PLUTOStatus = False Then
        StrSql = ""
        StrSql = "Select * from View_Clinic_MCO_ListingNew where 0 = 0"
        StrSql = StrSql + StrFormula
        rst.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
        Do While Not rst.EOF
            
            tmpBordxDate = Format(rst!MBMBordxDate, "yyyy/mm/dd")
            tmpPolEffDate = Format(rst!MBMPolEffDate, "yyyy/mm/dd")
            tmpPolExpDate = Format(rst!MBMPolEXPDate, "yyyy/mm/dd")
            tmpInvoiceDate = Format(rst!MBMCInvoiceDate, "yyyy/mm/dd")
            tmpMBMName = Replace(Trim(rst!MBMName), "'", "''")
            tmpMBMCName = Replace(Trim(rst!MBMCName), "'", "''")
            'tmpMBMBasicMCO = IIf(IsNumeric(rst!MBMBasicMCO), rst!MBMBasicMCO, 0)
            tmpMBMMCO = IIf(IsNumeric(rst!MBMCMCOFee), rst!MBMCMCOFee, 0)
            'tmpMBMMCOFee = IIf(IsNumeric(rst!MBMMCOFee), rst!MBMMCOFee, 0)
            'tmpPremium = IIf(IsNumeric(rst!MBMPremium), rst!MBMPremium, 0)
            If Len(rst!MBMDepartment) Then
                tmpDepartment = Replace(Trim(rst!MBMDepartment), "'", "''")
            Else
                tmpDepartment = ""
            End If
            If Len(rst!MBMGroupCompany) Then
                tmpGRPCompany = Replace(Trim(rst!MBMGroupCompany), "'", "''")
            Else
                tmpGRPCompany = ""
            End If
            
            
            sqlinsert = ""
            sqlinsert = "Insert into MCOInvoicePluto(MBMNumber,MBMBordxDate,MBMName,RELCode, " & _
                                                    "Type, MBMIcBcPp, MBMGroupCompany, MBMDepartment, HLTCode, " & _
                                                    "PAYCode, INSCode, MBMCLNStatus, MBMPolEffDate, MBMPolExpDate, " & _
                                                    "MBMCInvoiceNo,MBMCInvoiceDate, MBMCNumber, MBMCCoverID, MBMCName, " & _
                                                    "MBMCICBCPPNo, MBMCRELCode, MBMBatchNo, " & _
                                                    "GRPCoCode, MBMCMCOFee, MBMCInvoiced, MBMMemberType) " & _
                        "Values('" & rst!MBMNumber & "','" & tmpBordxDate & "', '" & tmpMBMName & "',  '" & rst!RELCODE & "'," & _
                        "'" & rst!Type & "', '" & rst!MBMIcBcPp & "', '" & tmpGRPCompany & "', '" & tmpDepartment & "', '" & rst!HLTCode & "'," & _
                        "'" & rst!PAYCode & "', '" & rst!INSCode & "', '" & rst!MBMCLNStatus & "', '" & tmpPolEffDate & "', '" & tmpPolExpDate & "', " & _
                        "'" & rst!MBMCInvoiceNo & "', '" & tmpInvoiceDate & "','" & rst!MBMCNUmber & "', '" & rst!MBMCCoverID & "', '" & tmpMBMCName & "', " & _
                        "'" & rst!MBMCICBCPPNo & "', '" & rst!MBMCRELCode & "','" & rst!MBMBatchNo & "', " & _
                        "'" & rst!GRPCoCode & "'," & tmpMBMMCO & ",'" & rst!MBMCInvoiced & "', '" & rst!MBMMemberType & "')"
            'MsgBox (sqlinsert)
            MediConnCISDBPluto.Execute sqlinsert
            
            
        rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
    End If

    If CboRptType.ListIndex = 0 Then
        If cboMCOListingVer = "MCO Listing - by Department" Then
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOInvoice\MCOInvoice_Dept.rpt"
        Else
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOInvoice\MCOInvoice.rpt"
        End If
    Else
        If cboMCOListingVer = "MCO Listing - by Department" Then
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOInvoice\MCOFeeListing_Dept.rpt"
        Else
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOInvoice\MCOFeeListing.rpt"
        End If
    End If
    Select Case cboMCOCriteriaSel.ListIndex ' By Batch
        Case 0
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(8)
            StrFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                    " {VIEW_Clinic_MCO_LISTINGNew.MBMBatchNo} = " & tmpBatchNo & ""
        Case 1
            StrFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") "
        Case 2
            StrFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMCInvoiceNo} = '" & Trim(txtInvoiceNo) & "'"
    End Select
    If cboMCOCriteriaSel.ListIndex <> 2 Then
        If CboPrintStatus.ListIndex = 0 Then
            StrFormula = StrFormula & " And {VIEW_Clinic_MCO_LISTINGNew.MBMCInvoiced} = 'N' "
        Else
            StrFormula = StrFormula & " And {VIEW_Clinic_MCO_LISTINGNew.MBMCInvoiced} = 'Y' "
        End If
        If cboMCOListingVer = "MCO Listing - by Payor" Then
            tmpPayCode = lvwMCO.SelectedItem
            StrFormula = StrFormula & " AND {VIEW_Clinic_MCO_LISTINGNew.PayCode} = '" & tmpPayCode & "'"
        Else
            tmpPayCode = Mid(cboMCOGrpCom, 1, 4)
            StrFormula = StrFormula & " AND  {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & tmpPayCode & "'"
        End If
    End If
    cRptMCO.SelectionFormula = StrFormula
    If cmdMCOPrintPreview Then
        cRptMCO.Destination = crptToWindow
        cRptMCO.WindowState = crptMaximized
    End If
    If cmdMCOPrint Then
        cRptMCO.Destination = crptToPrinter
    End If
    If CboRptType.ListIndex = 0 Then
        Call PrintMCOInvoice
    Else
        Call PrintMCOFeeListing
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PrintMCOFeeListing()
Const METHOD_NAME = "PrintMCOFeeListing"
On Error GoTo debugs


Dim tmpRptType As String
Dim tmpPayGrp As String
    cRptMCO.SortFields(0) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMNumber}"
    cRptMCO.SortFields(1) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMCCoverID}"
    Select Case cboMCOCriteriaSel.ListIndex
        Case 0 ' MCO Criteria Selection for Bordx Batch Print button
                tmpRptType = "BORDEREAUX DATE " & Space(9) & ":  " & txtMCOBordxDate & "       Batch No :    " & txtMCOBatchNo
        Case 1 ' MCO Criteia Selection for Bordx Range
                tmpRptType = "BORDEREAUX DATE " & Space(9) & ":  " & "FROM    " & txtMCOBordxDateFr & "      TO   " & txtMCOBordxDateTo
        Case 2 ' MCO Criteia Selection for Invoice No.
                tmpRptType = "INVOICE NO. " & Space(26) & ":  " & txtInvoiceNo
    End Select
    If cboMCOListingVer = "MCO Listing - by Payor" Then
        tmpPayGrp = "PAYOR " & Space(38) & ":  " & cboMCOPayor
    Else
        tmpPayGrp = "GROUP COMPANY " & Space(13) & ":  " & lvwMCO.SelectedItem.SubItems(10)
    End If
    cRptMCO.ParameterFields(0) = "PayorName;" & tmpPayGrp & ";true"
    cRptMCO.ParameterFields(1) = "RptType;" & tmpRptType & ";true"
    If cmdMCOPrint Then
        If bPrintPass = False And CancelPrinting = False Then
            bPrintPass = PrintFunction
            If bPrintPass = False Then
                Screen.MousePointer = Default
                Exit Sub
            End If
        End If
    End If
    cRptMCO.PrintReport
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmMCOFeeListNew")
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs
Dim PAY As New ADODB.Recordset
Dim sqlstr As String
Dim GrpCom As New ADODB.Recordset
        cboMCOPayor.Clear
        sqlstr = "select PAYCode, PAYCompanyName from PAY"
        PAY.Open sqlstr, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not PAY.EOF
            cboMCOPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
            PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
        cboMCOGrpCom.Clear
        sqlstr = "select GRPCoCode , GRPCoName from GRPCompany"
        GrpCom.Open sqlstr, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not GrpCom.EOF
            cboMCOGrpCom.AddItem Trim((GrpCom!GRPCoCode)) & " - " & Trim((GrpCom!GrpCoName))
            GrpCom.MoveNext
        Loop
        GrpCom.Close
        Set GrpCom = Nothing
        cboMCOListingVer.ListIndex = 1
        CboPrintStatus.ListIndex = 0
        CboRptType.ListIndex = 1
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function PrintFunction() As Boolean
On Error GoTo cancelPrint

    If cRptMCO Is Nothing Then Exit Function

        With dlgCommonDialog_MCO
            .DialogTitle = "Print Crystal Report "
            .CancelError = True
            .Flags = cdlPDNoSelection + cdlPDNoPageNums
            .ShowPrinter
            
            If Err <> MSComDlg.cdlCancel Then
                cRptMCO.PrintReport
            End If
        End With
        PrintFunction = True
        Exit Function
cancelPrint:
            PrintFunction = False
            CancelPrinting = True
            Exit Function
    
End Function

Private Sub PrintMCOInvoice()
Const METHOD_NAME = "PrintMCOInvoice"
On Error GoTo debugs

Dim AA As Integer
Dim tmpPayor As String
Dim tmpCompName As String
Dim tmpAdd1 As String
Dim tmpAdd2 As String
Dim tmpAdd3 As String
Dim tmpContact As String
Dim tmpTelNo As String
Dim tmpFaxNo As String
Dim InvoiceExist As Boolean
Dim MemberNo As String
Dim tmpMCO As Double
Dim CovID As String
Dim CompTrans As Boolean
Dim tmpInvNo As String
Dim tmpInvDate As String

    bPrintPass = False
    CancelPrinting = False
    CompTrans = False
    InvoiceExist = False
    tmpInvNo = ""
    For AA = 1 To lvwMCO.ListItems.Count
        If cboMCOListingVer.ListIndex = 0 Then
            tmpPayor = lvwMCO.ListItems(AA).Text
        Else
            tmpPayor = lvwMCO.ListItems(AA).SubItems(11)
        End If
        If Trim(lvwMCO.ListItems(AA).SubItems(14)) <> "" Then
            InvoiceExist = True
            tmpInvNo = Trim(lvwMCO.ListItems(AA).SubItems(14))
        End If
        tmpInvDate = IIf(Len(lvwMCO.ListItems(AA).SubItems(15)), lvwMCO.ListItems(AA).SubItems(15), "")
        Exit For
    Next AA
    
    Call PrintPAYInfo(tmpPayor, tmpCompName, tmpAdd1, tmpAdd2, tmpAdd3, tmpContact, tmpTelNo, tmpFaxNo)
    cRptMCO.SortFields(0) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate}"
    cRptMCO.SortFields(1) = "+{VIEW_Clinic_MCO_LISTINGNew.INSCode}"
    cRptMCO.ParameterFields(0) = "PayCode;" & tmpPayor & ";true"
    cRptMCO.ParameterFields(1) = "PayorName;" & tmpCompName & ";true"
    cRptMCO.ParameterFields(2) = "PayAdd1;" & tmpAdd1 & ";true"
    cRptMCO.ParameterFields(3) = "PayAdd2;" & tmpAdd2 & ";true"
    cRptMCO.ParameterFields(4) = "PayAdd3;" & tmpAdd3 & ";true"
    cRptMCO.ParameterFields(5) = "TelNo;" & tmpTelNo & ";true"
    cRptMCO.ParameterFields(6) = "FaxNo;" & tmpFaxNo & ";true"
    cRptMCO.ParameterFields(7) = "ContactName;" & tmpContact & ";true"
    cRptMCO.ParameterFields(8) = "InvoiceNo;" & tmpInvNo & ";true"
    cRptMCO.ParameterFields(9) = "InvoiceDate;" & tmpInvDate & ";true"

    If cmdMCOPrint Then
        If InvoiceExist = False Then
            On Err GoTo ErrUpdate
            medixmasterconn.beginTrans
            MediConnCISDB.beginTrans
            Call IncInvoiceNo
            cRptMCO.ParameterFields(8) = "InvoiceNo;" & txtInvoiceNo & ";true"
            cRptMCO.ParameterFields(9) = "InvoiceDate;" & Date & ";true"
            If bPrintPass = False And CancelPrinting = False Then
            bPrintPass = PrintFunction
                If bPrintPass = False Then
                    Screen.MousePointer = Default
                    Exit Sub
                Else
                    For AA = 1 To lvwMCO.ListItems.Count
                        If Trim(lvwMCO.ListItems(AA).ListSubItems(14).Text) = "" Then
                            lvwMCO.ListItems(AA).ListSubItems(13) = "YES"
                            lvwMCO.ListItems(AA).ListSubItems(14) = txtInvoiceNo
                            lvwMCO.ListItems(AA).ListSubItems(15) = Date
                            MemberNo = Trim(lvwMCO.ListItems(AA).ListSubItems(1))
                            CovID = Trim(lvwMCO.ListItems(AA).ListSubItems(2))
                            tmpMCO = CDbl(IIf(IsNumeric(lvwMCO.ListItems(AA).ListSubItems(7)), lvwMCO.ListItems(AA).ListSubItems(7), 0))
                            Call UpdateINVPrintingFile(MemberNo, CovID, tmpMCO)
                            cmdMCOPrint.Enabled = False
                        End If
                    Next AA
                   
                End If
                medixmasterconn.commitTrans
                MediConnCISDB.commitTrans
                CompTrans = True
            End If
        End If
    Else
        cRptMCO.PrintReport
    End If
    CompTrans = True
ErrUpdate:
    If CompTrans = False And InvoiceExist = True Then
        medixmasterconn.RollbackTrans
        MediConnCISDB.RollbackTrans
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub UpdateINVPrintingFile(tmpMBMNo As String, tmpCovID As String, tmpMCO As Double)
Const METHOD_NAME = "UpdateINVPrintingFile"
On Error GoTo debugs

Dim sqlUpd As String
Dim InvoiceDate As String

    InvoiceDate = Format(Date, "YYYY/MM/DD")
    If tmpCovID = "00" Then
        sqlUpd = " update MBMCLNOthers set " & _
            " MBMCInvoiceDate = '" & InvoiceDate & "', " & _
            " MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "', " & _
            " MBMCInvoiceBy = '" & Logon & "', MBMCInvoiced = 'Y', " & _
            " MBMCLNBilledMCO  = " & tmpMCO & _
            " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
            medixmasterconn.Execute sqlUpd
    Else
        sqlUpd = " update MBMCoveredPersonsCLN set " & _
            " MBMCInvoiceDate = '" & InvoiceDate & "', " & _
            " MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "', " & _
            " MBMCInvoiceBy = '" & Logon & "', " & _
            " MBMCCLNBilledMCO  = " & tmpMCO & _
            " Where MBMCNumber = '" & Trim(tmpMBMNo) & "'" & _
            " and MBMCCoverID = '" & tmpCovID & "'"
            medixmasterconn.Execute sqlUpd
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub IncInvoiceNo()
Const METHOD_NAME = "IncInvoiceNo"
On Error GoTo debugs

Dim MBMMCOINVQ As New ADODB.Recordset
Dim StrSql As String
Dim InvoiceNo As String
    StrSql = "Select * from CLNMCOINVQ "
    MBMMCOINVQ.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If MBMMCOINVQ.RecordCount = 0 Then
        With MBMMCOINVQ
            .AddNew
                !MCOInvoice_No = "00001"
                InvoiceNo = "00001"
            .update
        End With
    Else
        With MBMMCOINVQ
                !MCOInvoice_No = Format(!MCOInvoice_No + 1, "00000")
                InvoiceNo = Format(!MCOInvoice_No, "00000")
        .update
        End With
    End If
    txtInvoiceNo = "I" & InvoiceNo
    MBMMCOINVQ.Close
    Set MBMMCOINVQ = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtInvoiceNo_LostFocus()
    txtInvoiceNo = ChkStr(txtInvoiceNo)
End Sub

Private Sub txtMCOBordxDate_LostFocus()
    If txtMCOBordxDate <> "__/__/____" Then
        If IsDate(txtMCOBordxDate) = False Then
            MsgBox "Invalid Date format ! ", vbCritical + vbOKOnly
            txtMCOBordxDate.SetFocus
        End If
    End If
End Sub

Private Sub txtMCOBordxDateFr_LostFocus()
    If txtMCOBordxDateFr <> "__/__/____" Then
        If IsDate(txtMCOBordxDateFr) = False Then
            MsgBox "Invalid Date format ! ", vbCritical + vbOKOnly
            txtMCOBordxDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtMCOBordxDateTo_LostFocus()
    If txtMCOBordxDateTo <> "__/__/____" Then
        If IsDate(txtMCOBordxDateTo) = False Then
            MsgBox "Invalid Date format ! ", vbCritical + vbOKOnly
            txtMCOBordxDateTo.SetFocus
        End If
    End If
End Sub

Private Sub PrintPAYInfo(tmpPayor As String, tmpCompName As String, tmpAdd1 As String, tmpAdd2 As String, _
tmpAdd3 As String, tmpContact As String, tmpTelNo As String, tmpFaxNo As String)
Const METHOD_NAME = "PrintPAYInfo"
On Error GoTo debugs
Dim tmpPayGrp As New ADODB.Recordset
Dim strsqlPay As String
    If cboMCOListingVer = "Invoice - by Payor" Then
        strsqlPay = "SELECT PAYCompanyName AS CompanyName,PAYAddress1 as Add1, " & _
                    "PAYAddress2 as Add2,PAYAddress3 as Add3, PAYMAINContact as Contact, " & _
                    " PAYTELno as TelNo, PAYFAXNo as FaxNo " & _
                    " From PAY " & _
                    " WHERE PAYCODE = '" & tmpPayor & "'"
    Else
        strsqlPay = "SELECT GRPCoName AS CompanyName,GRPCoAdd1 as Add1, " & _
                    "GRPCoAdd2 as Add2,GRPCoAdd3 as Add3, " & _
                    "GRPCoContact as Contact, GRPCoPhone as TelNo, GRPCoFax as FaxNo " & _
                    " From GRPCompany " & _
                    " WHERE GRPCoCode = '" & tmpPayor & "'"
    End If
    tmpPayGrp.Open strsqlPay, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not tmpPayGrp.EOF
        tmpCompName = IIf(Len(tmpPayGrp!CompanyName), Trim(tmpPayGrp!CompanyName), "")
        tmpAdd1 = IIf(Len(tmpPayGrp!Add1), Trim(tmpPayGrp!Add1), "")
        tmpAdd2 = IIf(Len(tmpPayGrp!Add2), Trim(tmpPayGrp!Add2), "")
        tmpAdd3 = IIf(Len(tmpPayGrp!Add3), Trim(tmpPayGrp!Add3), "")
        tmpContact = IIf(Len(tmpPayGrp!contact), Trim(tmpPayGrp!contact), "")
        tmpTelNo = IIf(Len(tmpPayGrp!TelNo), Trim(tmpPayGrp!TelNo), "")
        tmpFaxNo = IIf(Len(tmpPayGrp!FaxNo), Trim(tmpPayGrp!FaxNo), "")
        tmpPayGrp.MoveNext
    Loop
    tmpPayGrp.Close
    Set tmpPayGrp = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub





