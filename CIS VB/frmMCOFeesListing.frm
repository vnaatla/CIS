VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMCOFeesListing 
   Caption         =   "Fee Listing"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Caption         =   "Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5400
      Left            =   120
      TabIndex        =   28
      Top             =   1800
      Width           =   11655
      Begin MSComctlLib.ListView lvwMCO 
         Height          =   5010
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8837
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Cover ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Relationship"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Bordereaux Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Effective Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Expiry Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Health Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Batch No"
            Object.Width           =   1764
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
            Text            =   "Printed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Department"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   7080
      Width           =   11655
      Begin VB.CommandButton cmdMCOClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9240
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrintPreview 
         Caption         =   "Pre&view"
         Height          =   300
         Left            =   8400
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   7560
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   300
         Left            =   6720
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label200 
         Alignment       =   2  'Center
         Caption         =   "Total Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label txtMCOTotalPrincipal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1320
         TabIndex        =   26
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label txtMCOTotalSupplementary1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         TabIndex        =   25
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label201 
         Alignment       =   2  'Center
         Caption         =   "Total Supp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Fee Listing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame FrameMCO 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         Left            =   7560
         TabIndex        =   31
         Top             =   930
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cboMCOCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCOFeesListing.frx":0000
         Left            =   2160
         List            =   "frmMCOFeesListing.frx":000A
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtMCOBordxDate 
         Height          =   315
         Left            =   7560
         TabIndex        =   2
         Top             =   285
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport cRptMCO 
         Left            =   9720
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         UserName        =   "sa"
         PrintFileLinesPerPage=   60
      End
      Begin MSMask.MaskEdBox txtMCOBordxDateTo 
         Height          =   315
         Left            =   7560
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMCOBordxDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   4
         Top             =   285
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_MCO 
         Left            =   9120
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtMCOBatchNo 
         Height          =   315
         Left            =   7560
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboMCOListingVer 
         Height          =   315
         ItemData        =   "frmMCOFeesListing.frx":0032
         Left            =   2160
         List            =   "frmMCOFeesListing.frx":003F
         TabIndex        =   5
         Top             =   650
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOPayor 
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   930
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOGrpCom 
         Height          =   315
         Left            =   2160
         TabIndex        =   30
         Top             =   930
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         Height          =   255
         Left            =   6600
         TabIndex        =   32
         Top             =   990
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl16 
         Caption         =   "Date "
         Height          =   255
         Left            =   6960
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   6600
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lbl15 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   6600
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbl18 
         Caption         =   "From"
         Height          =   255
         Left            =   6720
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl19 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbl22 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   680
         Width           =   1335
      End
      Begin VB.Label lbl20 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Single/Batch/Range"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmMCOFeesListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MCO_PAYCODE = 0
Const MCO_MEMBERNO = 1
Const MCO_COVERID = 2
Const MCO_RELATION = 3
Const MCO_BORDXDATE = 4
Const MCO_EFFDATE = 5
Const MCO_EXPDATE = 6
Const MCO_HLTCODE = 7
Const MCO_BATCHNO = 8
Const MCO_NAME = 9
Const MCO_GRPCO = 10
Const MCO_PRINTED = 11
Const MCO_DEPT = 12
Dim sDate As String
Dim DateFrom As String
Dim DateTo As String
Dim PAYCode As String
Dim BordxDat As String
Dim MBMNo As String
Dim Printed As String
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim TotalSupplementary As Integer

Private Sub cboMCOListingVer_Click()
Select Case cboMCOListingVer.ListIndex
    Case 0
        lbl20.Caption = "Payor"
        cboMCOPayor.Visible = True
        cboMCOGrpCom.Visible = False
        Label1.Visible = False
        CboDepartment.Visible = False
    Case 1
        lbl20.Caption = "Group Company"
        cboMCOGrpCom.Visible = True
        cboMCOPayor.Visible = False
        Label1.Visible = False
        CboDepartment.Visible = False
    Case 2
        lbl20.Caption = "Group Company"
        cboMCOGrpCom.Visible = True
        cboMCOPayor.Visible = False
        Label1.Visible = True
        CboDepartment.Visible = True
    End Select
End Sub

Private Sub cboMCOGrpCom_Click()
Dim sql As String
Dim GRPCOMName As String
Dim rst2 As New ADODB.Recordset

    If cboMCOListingVer.ListIndex = 2 Then
    
        sql = ""
        GRPCOMName = ""

        'GRPCOMName = Replace(cboMCOGrpCom, "'", "''")

        GRPCOMName = RemStr(Trim(Mid(cboMCOGrpCom, 7, IIf(InStr(7, cboMCOGrpCom, "-") = 0, 50, InStr(7, cboMCOGrpCom, "-") - 7))))
        
        '=Run Sp=================================================
        sql = "Exec SearchDepartment '" & Trim(GRPCOMName) & "'"
        rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        '===================================================
        'sql = " Select MBMDepartment From MBMOthers Where MBMGroupCompany = '" & Trim(GRPCOMName) & "' GROUP BY MBMDepartment"
        'rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    
        CboDepartment.Clear
        Do Until rst2.EOF
            With rst2
                CboDepartment.AddItem Trim(!MBMDepartment)
                rst2.MoveNext
            End With
        Loop
                
        rst2.Close
        Set rst2 = Nothing
    
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


Private Sub cboMCOCriteriaSel_Click()

cboMCOPayor = ""
CboDepartment = ""

Select Case cboMCOCriteriaSel.ListIndex
    Case 0
        lbl15.Visible = True
        lbl16.Visible = True
        lbl18.Visible = False
        Label4.Visible = True
        lbl19.Visible = False
        txtMCOBordxDate.Visible = True
        txtMCOBatchNo.Visible = True
        txtMCOBordxDateFr.Visible = False
        txtMCOBordxDateTo.Visible = False
        txtMCOBatchNo = ""
        cboMCOPayor = ""
        txtMCOBordxDate = "__/__/____"
        txtMCOBordxDateFr = "__/__/____"
        txtMCOBordxDateTo = "__/__/____"
        cboMCOListingVer.ListIndex = 0
        lvwMCO.ListItems.Clear
        cmdMCOPrint.Enabled = False
        cmdMCOPrintPreview.Enabled = False
        txtMCOTotalSupplementary1 = 0
        txtMCOTotalPrincipal = 0
    
    Case 1
        lbl15.Visible = False
        lbl16.Visible = False
        lbl18.Visible = True
        lbl19.Visible = True
        Label4.Visible = False
        txtMCOBordxDate.Visible = False
        txtMCOBatchNo.Visible = False
        txtMCOBordxDateFr.Visible = True
        txtMCOBordxDateTo.Visible = True
        cboMCOPayor = ""
        txtMCOBordxDate = "__/__/____"
        txtMCOBordxDateFr = "__/__/____"
        txtMCOBordxDateTo = "__/__/____"
        cboMCOListingVer.ListIndex = 0
        lvwMCO.ListItems.Clear
        cmdMCOPrint.Enabled = False
        cmdMCOPrintPreview.Enabled = False
        txtMCOTotalSupplementary1 = 0
        txtMCOTotalPrincipal = 0
    End Select

End Sub

Private Sub cboMCOListingVer_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboMCOListingVer.Text, cboMCOListingVer.SelStart) + Chr(KeyAscii) + Mid(cboMCOListingVer.Text, cboMCOListingVer.SelStart + cboMCOListingVer.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboMCOListingVer.ListCount - 1
        
        If NewText = LCase(Left(cboMCOListingVer.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMCOListingVer.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboMCOListingVer.Text = ValidValue
               cboMCOListingVer.SelStart = 0
               cboMCOListingVer.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboMCOPayor_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboMCOPayor.Text, cboMCOPayor.SelStart) + Chr(KeyAscii) + Mid(cboMCOPayor.Text, cboMCOPayor.SelStart + cboMCOPayor.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboMCOPayor.ListCount - 1
        
        If NewText = LCase(Left(cboMCOPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMCOPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboMCOPayor.Text = ValidValue
               cboMCOPayor.SelStart = 0
               cboMCOPayor.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cmdMCOClear_Click()
    txtMCOBatchNo = ""
    cboMCOPayor = ""
    txtMCOBordxDate = "__/__/____"
    txtMCOBordxDateFr = "__/__/____"
    txtMCOBordxDateTo = "__/__/____"
    cboMCOListingVer.ListIndex = 0
    lvwMCO.ListItems.Clear
    cmdMCOPrint.Enabled = False
    cmdMCOPrintPreview.Enabled = False
    txtMCOTotalSupplementary1 = 0
    txtMCOTotalPrincipal = 0
End Sub

Private Sub cmdMCOClose_Click()
    Unload Me
End Sub

Private Sub cmdMCOPrint_Click()
    MCOPrint "NO"
End Sub

Private Sub cmdMCOPrintPreview_Click()
    MCOPrint "NO"
End Sub

Private Sub cmdMCOSearch_Click()

txtMCOTotalPrincipal = ""
txtMCOTotalSupplementary1 = ""

    If cboMCOCriteriaSel = "" Then
        MsgBox "Please Enter MCO Criteria!", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    If cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 0 Then
        If txtMCOBordxDate = "__/__/____" Or txtMCOBatchNo = "" Then
            MsgBox "Please Enter Bordereaux Date and Batch No!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        ElseIf cboMCOPayor = "" Then
            MsgBox "Please Enter Payor!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    ElseIf cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 1 Then
        If txtMCOBordxDate = "__/__/____" Or txtMCOBatchNo = "" Then
            MsgBox "Please Enter Bordereaux Date and Batch No!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        ElseIf cboMCOGrpCom = "" Then
            MsgBox "Please Enter Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    ElseIf cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 2 Then
        If txtMCOBordxDate = "__/__/____" Or txtMCOBatchNo = "" Then
            MsgBox "Please Enter Bordereaux Date and Batch No!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        ElseIf cboMCOGrpCom = "" Then
            MsgBox "Please Enter Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        ElseIf CboDepartment = "" Then
            MsgBox "Please Enter Department!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 0 Then
        If txtMCOBordxDateFr = "__/__/____" Or txtMCOBordxDateTo = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Range!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        ElseIf cboMCOPayor = "" Then
            MsgBox "Please Enter Payor!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 1 Then
        If txtMCOBordxDateFr = "__/__/____" Or txtMCOBordxDateTo = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Range!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        ElseIf cboMCOGrpCom = "" Then
            MsgBox "Please Enter Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 2 Then
        If txtMCOBordxDateFr = "__/__/____" Or txtMCOBordxDateTo = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Range!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        ElseIf cboMCOGrpCom = "" Then
            MsgBox "Please Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        ElseIf CboDepartment = "" Then
            MsgBox "Please Enter Department!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    End If
    
    'If cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 0 And txtMCOBordxDate = "__/__/____" Or txtMCOBatchNo = "" Then
        'MsgBox "Please Enter Bordereaux Date and Batch No!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 0 And cboMCOPayor = "" Then
        'MsgBox "Please Enter Payor!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 1 And txtMCOBordxDate = "__/__/____" Or txtMCOBatchNo = "" Then
        'MsgBox "Please Enter Bordereaux Date and Batch No!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 1 And cboMCOGrpCom = "" Then
        'MsgBox "Please Enter Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 3 And txtMCOBordxDate = "__/__/____" Or txtMCOBatchNo = "" Then
        'MsgBox "Please Enter Bordereaux Date and Batch No!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 3 And cboMCOGrpCom = "" Then
        'MsgBox "Please Enter Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 0 And cboMCOListingVer.ListIndex = 3 And CboDepartment = "" Then
        'MsgBox "Please Enter Department!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 0 And txtMCOBordxDateFr = "__/__/____" Or txtMCOBordxDateTo = "__/__/____" Then
        'MsgBox "Please Enter Bordereaux Range!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 0 And cboMCOPayor = "" Then
        'MsgBox "Please Enter Payor!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 1 And txtMCOBordxDateFr = "__/__/____" Or txtMCOBordxDateTo = "__/__/____" Then
        'MsgBox "Please Enter Bordereaux Range!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 1 And cboMCOGrpCom = "" Then
        'MsgBox "Please Enter Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 2 And txtMCOBordxDateFr = "__/__/____" Or txtMCOBordxDateTo = "__/__/____" Then
        'MsgBox "Please Enter Bordereaux Range!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 2 And cboMCOGrpCom = "" Then
        'MsgBox "Please Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'ElseIf cboMCOCriteriaSel.ListIndex = 1 And cboMCOListingVer.ListIndex = 2 And CboDepartment = "" Then
        'MsgBox "Please Enter Department!", vbOKOnly + vbCritical, DialogBoxTitle
        'Exit Sub
    'End If
    
    If cboMCOListingVer = "" Then
        MsgBox "Please Enter Listing Version!", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    RecordSearching_MCO lvwMCO
    
End Sub

Private Sub RecordSearching_MCO(objListView As Object)

    Screen.MousePointer = vbHourglass
        Call MCOPrincipalListing
    Screen.MousePointer = Default
End Sub

Private Sub MCOPrincipalListing()
Dim rsSearch As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim sql As String
Dim itemToPrint As Integer
Dim isPrinted As Boolean
Dim CoverQty As Integer
Dim MCOCriteria As String
Dim TotalPrincipal As Integer
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String

TotalSupplementary = 0
Select Case cboMCOCriteriaSel.ListIndex
    Case 0 'Criteria Selection for MCO on Bordx Batch
            
            If cboMCOListingVer = "MCO Listing - by Payor" Then
                sDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
                lvwMCO.ListItems.Clear
                
                'sql = "SELECT * FROM VIEW_MBMClinic_MCO"
                'sql = sql & " WHERE MBMBordxDate = '" & sDate & "'"
                sql = sql & " AND MBMBordxDate = '" & sDate & "'"
                sql = sql & " AND MBMBatchNo = '" & txtMCOBatchNo & "'"
                sql = sql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
                
                sql = sql & " ORDER BY MBMNumber"
                'rsSearch.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic, adCmdText
                
                strAppendCase = "7"
            
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
                cmd.Parameters.Append prm1
                Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append prm2
                rsSearch.CursorLocation = adUseClient
                rsSearch.CursorType = adOpenForwardOnly
                rsSearch.CacheSize = 100
                Set rsSearch = cmd.Execute
                
            ElseIf cboMCOListingVer = "MCO Listing - by Group Company" Then
                sDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
                lvwMCO.ListItems.Clear
                
                'sql = "SELECT * FROM VIEW_MBMClinic_MCO_GrpCom"
                'sql = sql & " WHERE MBMBordxDate = '" & sDate & "'"
                sql = sql & " AND MBMBordxDate = '" & sDate & "'"
                sql = sql & " AND MBMBatchNo = '" & txtMCOBatchNo & "'"
                sql = sql & " AND GRPCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
                sql = sql & " ORDER BY MBMNumber"
                'rsSearch.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic, adCmdText
                
                strAppendCase = "8"
            
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
                cmd.Parameters.Append prm1
                Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append prm2
                rsSearch.CursorLocation = adUseClient
                rsSearch.CursorType = adOpenForwardOnly
                rsSearch.CacheSize = 100
                Set rsSearch = cmd.Execute
                
            
            ElseIf cboMCOListingVer = "MCO Listing - by Department" Then
                sDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
                lvwMCO.ListItems.Clear
                
                'sql = "SELECT * FROM VIEW_MBMClinic_MCO_GrpCom"
                'sql = sql & " WHERE MBMBordxDate = '" & sDate & "'"
                sql = sql & " AND MBMBordxDate = '" & sDate & "'"
                sql = sql & " AND MBMBatchNo = '" & txtMCOBatchNo & "'"
                sql = sql & " AND GRPCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
                sql = sql & " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
                sql = sql & " ORDER BY MBMNumber"
                'rsSearch.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic, adCmdText
                
                strAppendCase = "8"
            
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
                cmd.Parameters.Append prm1
                Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append prm2
                rsSearch.CursorLocation = adUseClient
                rsSearch.CursorType = adOpenForwardOnly
                rsSearch.CacheSize = 100
                Set rsSearch = cmd.Execute
                
            End If
            
    Case 1 'Criteria Selection for MCO on Bordx Range
            
            If cboMCOListingVer = "MCO Listing - by Payor" Then
                DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
                DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
                lvwMCO.ListItems.Clear
                
                'sql = "SELECT * FROM VIEW_MBMClinic_MCO "
                'sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                sql = sql & " AND MBMBordxDate >= '" & DateFrom & "'"
                sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
                sql = sql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
                sql = sql & " ORDER BY MBMNumber"
                'rsSearch.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic, adCmdText
                
                strAppendCase = "7"
            
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
                cmd.Parameters.Append prm1
                Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append prm2
                rsSearch.CursorLocation = adUseClient
                rsSearch.CursorType = adOpenForwardOnly
                rsSearch.CacheSize = 100
                Set rsSearch = cmd.Execute
                
            
            ElseIf cboMCOListingVer = "MCO Listing - by Group Company" Then
                DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
                DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
                lvwMCO.ListItems.Clear
                
                'sql = "SELECT * FROM VIEW_MBMClinic_MCO_GrpCom "
                'sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                sql = sql & " AND MBMBordxDate >= '" & DateFrom & "'"
                sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
                sql = sql & " AND GRPCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
                sql = sql & " ORDER BY MBMNumber"
                'rsSearch.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic, adCmdText
                
                strAppendCase = "8"
            
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
                cmd.Parameters.Append prm1
                Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append prm2
                rsSearch.CursorLocation = adUseClient
                rsSearch.CursorType = adOpenForwardOnly
                rsSearch.CacheSize = 100
                Set rsSearch = cmd.Execute
            
            ElseIf cboMCOListingVer = "MCO Listing - by Department" Then
                DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
                DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
                lvwMCO.ListItems.Clear
                
                'sql = "SELECT * FROM VIEW_MBMClinic_MCO_GrpCom "
                'sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                sql = sql & " AND MBMBordxDate >= '" & DateFrom & "'"
                sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
                sql = sql & " AND GRPCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
                sql = sql & " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
                sql = sql & " ORDER BY MBMNumber"
                'rsSearch.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic, adCmdText
                
                strAppendCase = "8"
            
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "MCOFeeModule"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
                cmd.Parameters.Append prm1
                Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
                cmd.Parameters.Append prm2
                rsSearch.CursorLocation = adUseClient
                rsSearch.CursorType = adOpenForwardOnly
                rsSearch.CacheSize = 100
                Set rsSearch = cmd.Execute
            
            End If
            
End Select


    'If rsSearch.RecordCount = 0 Then
        'MsgBox "Record not found !", vbInformation + vbOKOnly
        'cmdMCOPrint.Enabled = False
        'Screen.MousePointer = Default
        'Exit Sub
    'End If
              
    
    Do Until rsSearch.EOF
    Total = rsSearch.RecordCount
    lblTotal = CInt(Total)
        Set ListMaster = lvwMCO.ListItems.Add(, , rsSearch!PAYCode)
            ListMaster.SubItems(1) = rsSearch!MBMNumber
            ListMaster.SubItems(2) = "00"
            ListMaster.SubItems(3) = rsSearch!MBMMemberType
            ListMaster.SubItems(4) = rsSearch!MBMBordxDate
            ListMaster.SubItems(5) = rsSearch!MBMPolEffDate
            ListMaster.SubItems(6) = rsSearch!MBMPolExpDate
            'If Len(rsSearch!MBMPolicyNo) Then
                'ListMaster.SubItems(5) = Trim(rsSearch!MBMPolicyNo)
            'End If
            ListMaster.SubItems(7) = Trim(rsSearch!HLTCode)
            If Len(rsSearch!MBMBatchNo) Then
                ListMaster.SubItems(8) = rsSearch!MBMBatchNo
            End If
            ListMaster.SubItems(9) = Trim(rsSearch!MBMName)
                                
            If Len(Trim(rsSearch!MBMGroupCompany)) = 0 Or IsNull(rsSearch!MBMGroupCompany) = True Then
                ListMaster.SubItems(10) = ""
            Else
                ListMaster.SubItems(10) = Trim(rsSearch!MBMGroupCompany)
            End If
            
            If Len(Trim(rsSearch!MBMDepartment)) = 0 Or IsNull(rsSearch!MBMDepartment) = True Then
                ListMaster.SubItems(12) = ""
            Else
                ListMaster.SubItems(12) = Trim(rsSearch!MBMDepartment)
            End If
                            
                TotalPrincipal = TotalPrincipal + 1
                txtMCOTotalPrincipal = TotalPrincipal
                PAYCode = rsSearch!PAYCode
                MBMNo = rsSearch!MBMNumber
                BordxDat = rsSearch!MBMBordxDate
                
                'itemToPrint = MCO_PRINTED
                cmdMCOPrint.Enabled = True
                cmdMCOPrintPreview.Enabled = True
                                                
                If Trim(rsSearch!INSCode) = "F" Or Trim(rsSearch!INSCode) = "H" Then
                    Call MCOSuppListing
                End If
                    
    rsSearch.MoveNext
    Loop
    
    If lvwMCO.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdMCOPrint.Enabled = False
        Screen.MousePointer = Default
    End If
    
    rsSearch.Close
    Set rsSearch = Nothing

End Sub

Private Sub MCOSuppListing()
Dim strsql As String
Dim sql As String
Dim strsqls As String
Dim rsPrint As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim ListMaster As ListItem
Dim isPrinted As Boolean
Dim Total As Integer
Dim Current As Integer
Dim cmd2 As New ADODB.Command
Dim prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim strAppendCase As String


    If cboMCOListingVer = "MCO Listing - by Payor" Then
        Select Case cboMCOCriteriaSel.ListIndex
            Case 0
                    sDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
                    'strsql = "SELECT * FROM VIEW_MBMClinic_MCO_Supp "
                    'strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "'"
                    strsql = strsql & " AND MBMBordxDate = '" & sDate & "'"
                    strsql = strsql & " AND MBMBatchNo = '" & Trim(txtMCOBatchNo) & "'"
                    strsql = strsql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            Case 1
                    DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
                    DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
                    'strsql = "SELECT * FROM VIEW_MBMClinic_MCO_Supp "
                    'strsql = strsql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate <= '" & DateTo & "'"
                    strsql = strsql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        End Select
    
    ElseIf cboMCOListingVer = "MCO Listing - by Group Company" Then
    
        Select Case cboMCOCriteriaSel.ListIndex
        Case 0
                sDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
                'strsql = "SELECT * FROM VIEW_MBMClinic_MCO_Supp "
                'strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "'"
                strsql = strsql & " AND MBMBordxDate = '" & sDate & "'"
                strsql = strsql & " AND MBMBatchNo = '" & Trim(txtMCOBatchNo) & "'"
                strsql = strsql & " AND GrpCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
                strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                
                'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        Case 1
                DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
                DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
                'strsql = "SELECT * FROM VIEW_MBMClinic_MCO_Supp "
                'strsql = strsql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                strsql = strsql & " AND MBMBordxDate >= '" & DateFrom & "'"
                strsql = strsql & " AND MBMBordxDate <= '" & DateTo & "'"
                strsql = strsql & " AND GrpCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
                strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                
                'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        End Select
        
    ElseIf cboMCOListingVer = "MCO Listing - by Department" Then
    
        Select Case cboMCOCriteriaSel.ListIndex
        Case 0
                sDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
                'strsql = "SELECT * FROM VIEW_MBMClinic_MCO_Supp "
                'strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "'"
                strsql = strsql & " AND MBMBordxDate = '" & sDate & "'"
                strsql = strsql & " AND MBMBatchNo = '" & Trim(txtMCOBatchNo) & "'"
                strsql = strsql & " AND GrpCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
                strsql = strsql & " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
                strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                
                'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        Case 1
                DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
                DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
                'strsql = "SELECT * FROM VIEW_MBMClinic_MCO_Supp "
                'strsql = strsql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                strsql = strsql & " AND MBMBordxDate >= '" & DateFrom & "'"
                strsql = strsql & " AND MBMBordxDate <= '" & DateTo & "'"
                strsql = strsql & " AND GrpCoCode = '" & Mid(cboMCOGrpCom, 1, 4) & "'"
                strsql = strsql & " AND MBMDepartment = '" & Trim(CboDepartment) & "'"
                strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                
                'MBMCov.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        End Select
        
    End If
    
    strAppendCase = "9"
            
    Set cmd2.ActiveConnection = MediConnCISDB
    cmd2.CommandText = "MCOFeeModule"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 300
    Set prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
    cmd2.Parameters.Append prm3
    Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd2.Parameters.Append prm4
    MBMCov.CursorLocation = adUseClient
    MBMCov.CursorType = adOpenForwardOnly
    MBMCov.CacheSize = 100
    Set MBMCov = cmd2.Execute

        
   Do Until MBMCov.EOF
       Set ListMaster = lvwMCO.ListItems.Add(, , PAYCode)
            ListMaster.SubItems(1) = MBMCov!MBMCNUmber
            ListMaster.SubItems(2) = MBMCov!MBMCCoverID
            ListMaster.SubItems(3) = MBMCov!MBMCRELCode
            ListMaster.SubItems(4) = lvwMCO.SelectedItem.SubItems(4)
            ListMaster.SubItems(5) = Trim(lvwMCO.SelectedItem.SubItems(5))
            ListMaster.SubItems(6) = Trim(lvwMCO.SelectedItem.SubItems(6))
            ListMaster.SubItems(7) = lvwMCO.SelectedItem.SubItems(7)
            ListMaster.SubItems(8) = lvwMCO.SelectedItem.SubItems(8)
            ListMaster.SubItems(9) = Trim(MBMCov!MBMCName)
        
           If MBMCov!MBMGroupCompany = "" Then
                ListMaster.SubItems(10) = ""
            Else
                If Len(MBMCov!MBMGroupCompany) Then
                    ListMaster.SubItems(10) = Trim(MBMCov!MBMGroupCompany)
                End If
            End If
        
           TotalSupplementary = TotalSupplementary + 1
           txtMCOTotalSupplementary1 = TotalSupplementary
           sql = " SELECT MBMNumber, PAYCode, HLTCode, MBMBordxDate, " & _
                 " MBMBatchNo FROM MBM "
           BordxDat = CStr(Month(MBMCov!MBMBordxDate)) & "/" & CStr(Day(MBMCov!MBMBordxDate)) & "/" & CStr(Year(MBMCov!MBMBordxDate))
      
           sql = sql & " WHERE MBMBordxDate = '" & BordxDat & "'"
           sql = sql & " AND PAYCode = '" & PAYCode & "'"
           sql = sql & " AND HLTCode = '" & lvwMCO.SelectedItem.SubItems(7) & "'"
           sql = sql & " AND MBMBatchno = '" & lvwMCO.SelectedItem.SubItems(8) & "'"
           sql = sql & " AND MBMNumber = '" & MBMCov!MBMCNUmber & "'"
           
           rsPrint.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
           
            itemToPrint = MCO_PRINTED
            cmdMCOPrint.Enabled = True
            cmdMCOPrintPreview.Enabled = True
           
           If Printed = "YES" Then
               ListMaster.SubItems(itemToPrint) = "YES"
           Else
               ListMaster.SubItems(itemToPrint) = "NO"
           End If
           
           rsPrint.Close
           MBMCov.MoveNext
       Loop

    MBMCov.Close
    Set rsPrint = Nothing
    Set MBMCov = Nothing

End Sub

Private Sub MCOPrint(PrintMode As String)
Dim ii As Integer
Dim MemberNo As String

Dim tmpBDate, tmpBatchNo, TMPPayCode, tmpHLTCode, tmpMemberNo, TMPDepartment As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tempPrintedHLTcode, tmpPrintedMemberNo, tmpPrintedDepartment As String
Dim BDateArray() As String
Dim BatchNoArray() As String
Dim PaycodeArray() As String
Dim HLTcodeArray() As String
Dim MemberNoArray() As String
Dim DepartmentArray() As String

Dim Cnt As Integer
Dim cont As Integer
Dim cout As Integer
Dim Arraycnt As Integer
Dim Arraycount As Integer
Dim PrintM As String
Dim i As Integer
Dim A As Integer
Dim B As Integer
Dim rstInv As New ADODB.Recordset
Dim strsql2 As String
Dim InvNoFr, InvNoTo As String

Screen.MousePointer = vbHourglass

Select Case cboMCOCriteriaSel.ListIndex
    Case 0 'MCO Criteria Selection for Bordx Batch
            Cnt = 0
            
            For ii = 1 To lvwMCO.ListItems.Count
                 Set lvwMCO.SelectedItem = lvwMCO.ListItems(ii)
                 If lvwMCO.SelectedItem.SubItems(MCO_PRINTED) = PrintMode Then
                     tmpMemberNo = lvwMCO.SelectedItem.SubItems(MCO_MEMBERNO)
                     tmpBDate = lvwMCO.SelectedItem.SubItems(MCO_BORDXDATE)
                     tmpHLTCode = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                     tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                     TMPPayCode = lvwMCO.SelectedItem
                     
                     If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And TMPPayCode = tmpPrintedPaycode And _
                         tmpHLTCode = tempPrintedHLTcode And tmpMemberNo = tmpPrintedMemberNo Then
                     Else
                         tmpPrintedBDate = tmpBDate
                         tmpPrintedBatchNo = tmpBatchNo
                         tmpPrintedPaycode = TMPPayCode
                         tempPrintedHLTcode = tmpHLTCode
                         tmpPrintedMemberNo = tmpMemberNo
                         Cnt = Cnt + 1
                         ReDim Preserve BDateArray(Cnt)
                         ReDim Preserve BatchNoArray(Cnt)
                         ReDim Preserve PaycodeArray(Cnt)
                         ReDim Preserve HLTcodeArray(Cnt)
                         ReDim Preserve MemberNoArray(Cnt)
                         BDateArray(Cnt) = tmpBDate
                         BatchNoArray(Cnt) = tmpBatchNo
                         PaycodeArray(Cnt) = TMPPayCode
                         HLTcodeArray(Cnt) = tmpHLTCode
                         MemberNoArray(Cnt) = tmpMemberNo
                     End If
                 End If
            Next ii
            
    
    Case 1 'MCO Criteria Selection for Bordx Range
            Cnt = 0
            
                For ii = 1 To lvwMCO.ListItems.Count
                      Set lvwMCO.SelectedItem = lvwMCO.ListItems(ii)
                      If lvwMCO.SelectedItem.SubItems(MCO_PRINTED) = PrintMode Then
                          tmpMemberNo = lvwMCO.SelectedItem.SubItems(MCO_MEMBERNO)
                          tmpBDate = lvwMCO.SelectedItem.SubItems(MCO_BORDXDATE)
                          tmpHLTCode = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                          tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                          TMPPayCode = lvwMCO.SelectedItem
                          
                          If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And TMPPayCode = tmpPrintedPaycode And _
                              tmpHLTCode = tempPrintedHLTcode And tmpMemberNo = tmpPrintedMemberNo Then
                          Else
                              tmpPrintedBDate = tmpBDate
                              tmpPrintedBatchNo = tmpBatchNo
                              tmpPrintedPaycode = TMPPayCode
                              tempPrintedHLTcode = tmpHLTCode
                              tmpPrintedMemberNo = tmpMemberNo
                              Cnt = Cnt + 1
                              ReDim Preserve BDateArray(Cnt)
                              ReDim Preserve BatchNoArray(Cnt)
                              ReDim Preserve PaycodeArray(Cnt)
                              ReDim Preserve HLTcodeArray(Cnt)
                              ReDim Preserve MemberNoArray(Cnt)
                              BDateArray(Cnt) = tmpBDate
                              BatchNoArray(Cnt) = tmpBatchNo
                              PaycodeArray(Cnt) = TMPPayCode
                              HLTcodeArray(Cnt) = tmpHLTCode
                              MemberNoArray(Cnt) = tmpMemberNo
                          End If
                      End If
                 Next ii

End Select


If cmdMCOPrintPreview Then
    Call MCOListingVer
Else
    If cmdMCOPrint Then
            A = 0
            B = 0
            For i = 1 To lvwMCO.ListItems.Count
                Set lvwMCO.SelectedItem = lvwMCO.ListItems(i)
                    PrintM = lvwMCO.SelectedItem.SubItems(11)
                    If PrintM = "YES" Then
                        A = A + 1
                    Else
                        B = B + 1
                    End If
            Next i
        
        Select Case cboMCOCriteriaSel.ListIndex
            Case 0 'MCO Criteria Selection for Bordx Batch Print
                Call MCOListingVer
            
            Case 1 'MCO Criteria Selection for Bordx Range Print
                Call MCOListingVer
        
        End Select
    End If
End If
Screen.MousePointer = Default
End Sub

Private Sub MCOListingVer()
Dim i As Integer
Dim rsSearch As New ADODB.Recordset

bPrintPass = False
CancelPrinting = False
If cmdMCOPrintPreview Then
    Select Case cboMCOCriteriaSel.ListIndex
        Case 0 ' MCO Criteria Selection for Bordx Batch Print Preview button
                Call MCOListingVer_BBPrintPrev
            
        Case 1 ' MCO Criteria Selection for Bordx Range Print Preview
                Call MCOListingVer_BRPrintPreV
    End Select
Else
    If cmdMCOPrint Then
        Select Case cboMCOCriteriaSel.ListIndex
            Case 0 'MCO Criteria Selection for Bordx Batch Print button
                    Call MCOListingVer_BBPrint
            
            Case 1 ' MCO Criteia Selection for Bordx Range
                    Call MCOListingVer_BRPrint
           
        End Select
    End If
End If

End Sub

Private Sub MCOListingVer_BBPrintPrev()

    Select Case cboMCOListingVer
        Case "MCO Listing - by Payor"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = lvwMCO.SelectedItem
            sDate = Trim(txtMCOBordxDate)
                                 
            ss = "{VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremiumNew.rpt"
            
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.PayCode} = '" & Mid(cboMCOPayor, 1, 2) & "'" & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBatchNo} = " & txtMCOBatchNo & ""
                                                    
            Call MCOListingBBPrintSeqPrev
                     
        Case "MCO Listing - by Group Company"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = Mid(cboMCOGrpCom, 1, 4)
            sDate = Trim(txtMCOBordxDate)
                                 
            ss = "{VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremium_GrpComNew.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & TMPPayCode & "'" & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBatchNo} = " & tmpBatchNo & ""
                         
            Call MCOListingBBPrintSeqPrev
            
        Case "MCO Listing - by Department"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = Mid(cboMCOGrpCom, 1, 4)
            TMPDepartment = Trim(CboDepartment)
            sDate = Trim(txtMCOBordxDate)
                                 
            ss = "{VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremium_Dept.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & TMPPayCode & "'" & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMDepartment} = '" & TMPDepartment & "'" & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBatchNo} = " & tmpBatchNo & ""
                         
            Call MCOListingBBPrintSeqPrev
                         
    End Select

End Sub

Private Sub MCOListingVer_BRPrintPreV()

    Select Case cboMCOListingVer
        Case "MCO Listing - by Payor"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = lvwMCO.SelectedItem
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
                                    
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremiumNew.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.PayCode} = '" & TMPPayCode & "'"
                                                       
            Call MCOListingBRPrintSeqPrev
                                
        Case "MCO Listing - by Group Company"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = Mid(cboMCOGrpCom, 1, 4)
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
            
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremium_GrpComNew.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & TMPPayCode & "'"
                            
           Call MCOListingBRPrintSeqPrev
           
        Case "MCO Listing - by Department"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = Mid(cboMCOGrpCom, 1, 4)
            TMPDepartment = Trim(CboDepartment)
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
            
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremium_Dept.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & TMPPayCode & "' AND {VIEW_Clinic_MCO_LISTINGNew.MBMDepartment} = '" & TMPDepartment & "'"
                            
           Call MCOListingBRPrintSeqPrev
                            
    End Select
    
End Sub

Private Sub MCOListingVer_BBPrint()

    Select Case cboMCOListingVer
        Case "MCO Listing - by Payor"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = lvwMCO.SelectedItem
            sDate = Trim(txtMCOBordxDate)
                                        
            ss = "{VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremiumNew.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBatchNo} = " & tmpBatchNo & ""
                                
            Call MCOListingBBPrintSeq
                                    
        Case "MCO Listing - by Group Company"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = Mid(cboMCOGrpCom, 1, 4)
            sDate = Trim(txtMCOBordxDate)
                                        
            ss = "{VIEW_Clinic_MCO_LISTINGNMew.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremium_GrpComNew.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & TMPPayCode & "'" & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBatchNo} = " & tmpBatchNo & ""
                                       
            Call MCOListingBBPrintSeq
        
        Case "MCO Listing - by Department"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = Mid(cboMCOGrpCom, 1, 4)
            TMPDepartment = Trim(CboDepartment)
            sDate = Trim(txtMCOBordxDate)
                                 
            ss = "{VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremium_Dept.rpt"
            
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & TMPPayCode & "'" & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMDepartment} = '" & TMPDepartment & "'" & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBatchNo} = " & tmpBatchNo & ""
                         
            Call MCOListingBBPrintSeq
                                    
    End Select
    
End Sub

Private Sub MCOListingVer_BRPrint()

    Select Case cboMCOListingVer
        Case "MCO Listing - by Payor"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = lvwMCO.SelectedItem
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
                                        
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremiumNew.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.PayCode} = '" & TMPPayCode & "'"
                                                       
            Call MCOListingBRPrintSeq
                                    
        Case "MCO Listing - by Group Company"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = Mid(cboMCOGrpCom, 1, 4)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
                            
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremium_GrpComNew.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & TMPPayCode & "'"
                                                      
            Call MCOListingBRPrintSeq
            
        Case "MCO Listing - by Department"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = Mid(cboMCOGrpCom, 1, 4)
            TMPDepartment = Trim(CboDepartment)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
                            
            cRptMCO.ReportFileName = crdPath & "\CISReport\MCOListing_WithoutPremium_Dept.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_Clinic_MCO_LISTINGNew.GrpCoCode} = '" & TMPPayCode & "' AND {VIEW_Clinic_MCO_LISTINGNew.MBMDepartment} = '" & TMPDepartment & "'"
                                                      
            Call MCOListingBRPrintSeq
                                
    End Select

End Sub

Private Sub MCOListingBBPrintSeqPrev()

    cRptMCO.SortFields(0) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMNumber}"
    cRptMCO.SortFields(1) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
    cRptMCO.ParameterFields(1) = "FromDate;" & sDate & ";true"
    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
    If cboMCOListingVer = "MCO Listing - by Payor" Then
        cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
    Else
        cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOGrpCom & ";true"
    End If
    cRptMCO.PrintReport

End Sub

Private Sub MCOListingBRPrintSeqPrev()
                    
    cRptMCO.SortFields(0) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMNumber}"
    cRptMCO.SortFields(1) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
    If cboMCOListingVer = "MCO Listing - by Payor" Then
        cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
    Else
        cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOGrpCom & ";true"
    End If
    cRptMCO.PrintReport
                
End Sub

Private Sub MCOListingBBPrintSeq()

    cRptMCO.SortFields(0) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMNumber}"
    cRptMCO.SortFields(1) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
    cRptMCO.ParameterFields(1) = "FromDate;" & sDate & ";true"
    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
    
    If cboMCOListingVer = "MCO Listing - by Payor" Then
        cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
    Else
        cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOGrpCom & ";true"
    End If
    
    If bPrintPass = False And CancelPrinting = False Then
        bPrintPass = PrintFunction
        If bPrintPass = False Then
            Screen.MousePointer = Default
            Exit Sub
        End If
    Else
        cRptMCO.PrintReport
    End If
                    
End Sub

Private Sub MCOListingBRPrintSeq()

    cRptMCO.SortFields(0) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMNumber}"
    cRptMCO.SortFields(1) = "+{VIEW_Clinic_MCO_LISTINGNew.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
    
    If cboMCOListingVer = "MCO Listing - by Payor" Then
        cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
    Else
        cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOGrpCom & ";true"
    End If
    
    If bPrintPass = False And CancelPrinting = False Then
        bPrintPass = PrintFunction
        If bPrintPass = False Then
            Screen.MousePointer = Default
            Exit Sub
        End If
    Else
        cRptMCO.PrintReport
    End If
                
End Sub

Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim sqlstr As String
Dim GRPCOM As New ADODB.Recordset
Dim sqlstr2 As String
   
    sqlstr = "select PAYCode, PAYCompanyName from PAY"
    
    PAY.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    Do While Not PAY.EOF
    
        cboMCOPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    
    Loop
    
    PAY.Close
    Set PAY = Nothing
    
    sqlstr = "select GRPCoCode , GRPCoName from GRPCompany"
    
    GRPCOM.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    Do While Not GRPCOM.EOF
    
        cboMCOGrpCom.AddItem Trim((GRPCOM!GRPCoCode)) & " - " & Trim((GRPCOM!GrpCoName))
        GRPCOM.MoveNext
    
    Loop
    
    GRPCOM.Close
    Set GRPCOM = Nothing
    
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
 
Private Sub Form_Load()
    Call ComboListing
    'Call FormTabView
       X = CreatePath("c:\cardPrinting")
        If X = False Then
            MsgBox "Error Creating Folder c:\CardPrinting !" & vbNewLine & _
                    " Please create that folder manually to store the output of leaflet in sofe copy!"
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
