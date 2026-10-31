VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCOInvoiceNew 
   Caption         =   "MCO Invoice"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
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
      Height          =   5160
      Left            =   120
      TabIndex        =   30
      Top             =   2160
      Width           =   11055
      Begin MSComctlLib.ListView lvwInvoice 
         Height          =   4755
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   10815
         _ExtentX        =   19071
         _ExtentY        =   8382
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor Code"
            Object.Width           =   2540
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
            Object.Width           =   2117
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
            SubItemIndex    =   5
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Health Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Batch No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Printed"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Invoice No"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   7200
      Width           =   11055
      Begin VB.TextBox lblTotSupp 
         Height          =   285
         Left            =   3840
         TabIndex        =   33
         Text            =   "0"
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox lblTotPrinc 
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Text            =   "0"
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton cmdINVClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9240
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrintPreview 
         Caption         =   "Pre&view"
         Enabled         =   0   'False
         Height          =   300
         Left            =   8400
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7560
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   300
         Left            =   6720
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   285
         Left            =   5880
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdINVClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label203 
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
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label202 
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
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MCO Invoice"
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
         TabIndex        =   20
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame FrameINV 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   11055
      Begin VB.TextBox txtINVBatchNo 
         Height          =   315
         Left            =   8520
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtINVBordxDate 
         Height          =   315
         Left            =   6000
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport cRptINV 
         Left            =   9720
         Top             =   1680
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         UserName        =   "sa"
         PrintFileLinesPerPage=   60
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSMask.MaskEdBox txtINVBordxDateTo 
         Height          =   315
         Left            =   8520
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtINVBordxDateFr 
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_INV 
         Left            =   8880
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtINVInvoiceNo 
         Height          =   315
         Left            =   6000
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboINVCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCOInvoiceNew.frx":0000
         Left            =   2160
         List            =   "frmMCOInvoiceNew.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cboINVListingVer 
         Height          =   315
         ItemData        =   "frmMCOInvoiceNew.frx":0041
         Left            =   2160
         List            =   "frmMCOInvoiceNew.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   650
         Width           =   5415
      End
      Begin VB.ComboBox cboINVPayor 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   930
         Width           =   5415
      End
      Begin VB.ComboBox cboINVGrpCom 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   930
         Width           =   5415
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cboINVPrintOpt 
         Height          =   315
         ItemData        =   "frmMCOInvoiceNew.frx":009B
         Left            =   2160
         List            =   "frmMCOInvoiceNew.frx":00A5
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lbl29 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         Height          =   255
         Left            =   5040
         TabIndex        =   17
         Top             =   1260
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl30 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   7560
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl36 
         Caption         =   "Print Option"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl35 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lbl34 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lbl27 
         Caption         =   "Single/Batch/Range"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmMCOInvoiceNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SuppIDList() As String
Dim tmpTotSupp As Integer
Dim GetSubSupp As String
Dim GetSuppMCO As String
Dim MultiInvFound As Boolean
Dim PrevInvNO As String
Dim CurrInvNo As String

Private Sub cboINVGrpCom_LostFocus()
Dim sql As String
Dim GRPCOMName As String
Dim rst2 As New ADODB.Recordset

    If cboINVListingVer.ListIndex = 2 Then
    
        sql = ""
        GRPCOMName = ""
        GRPCOMName = RemStr(Trim(Mid(cboINVGrpCom, 7, IIf(InStr(7, cboINVGrpCom, "-") = 0, 50, InStr(7, cboINVGrpCom, "-") - 7))))
        
        '=Run Sp=================================================
        sql = "Exec SearchDepartment '" & Trim(GRPCOMName) & "'"
        rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        '===================================================
        'sql = " Select MBMDepartment From MBMOthers Where MBMGroupCompany = '" & Trim(GRPCOMName) & "' GROUP BY MBMDepartment"
    
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

Private Sub RecordSearching_INV(objListView As Object)
    Screen.MousePointer = vbHourglass
    
        Call INVPrincipalListing
    
    Screen.MousePointer = Default
End Sub

Private Sub INVPrincipalListing()
Const METHOD_NAME = "INVPrincipalListing"
On Error GoTo debugs

Dim rsSearch As New ADODB.Recordset
Dim tmpTotPrinc As Integer
Dim AA As Integer
Dim tmpSuppID As String
Dim ListMaster As ListItem
Dim suppExist As Boolean

    tmpTotSupp = 0
    lvwInvoice.ListItems.Clear
    MultiInvFound = False
    ReDim SuppIDList(0)
    txtInvoiceNo = ""
    Call ReadRecord(rsSearch, 1, "P", False, True, False, "")
        
        Do Until rsSearch.EOF
            Set ListMaster = lvwInvoice.ListItems.Add(, , rsSearch!PAYCode)
            ListMaster.SubItems(1) = Trim(rsSearch!MBMNumber)
            ListMaster.SubItems(2) = "00"
            ListMaster.SubItems(3) = rsSearch!MBMMemberType
            ListMaster.SubItems(4) = rsSearch!MBMBordxDate
            If Len(rsSearch!MBMPolicyNo) Then
                ListMaster.SubItems(5) = Trim(rsSearch!MBMPolicyNo)
            End If
            If Len(rsSearch!MBMBatchNo) Then
                ListMaster.SubItems(7) = rsSearch!MBMBatchNo
            End If
            ListMaster.SubItems(8) = Trim(rsSearch!MBMName)
                                
            If Len(Trim(rsSearch!MBMGroupCompany)) = 0 Or IsNull(rsSearch!MBMGroupCompany) = True Then
                ListMaster.SubItems(9) = ""
            Else
                ListMaster.SubItems(9) = Trim(rsSearch!MBMGroupCompany)
            End If
                 
            If Len(rsSearch!MBMCInvoiceNo) Then
                txtInvoiceNo = Trim(rsSearch!MBMCInvoiceNo)
                ListMaster.SubItems(11) = Trim(rsSearch!MBMCInvoiceNo)
                CurrInvNo = Trim(rsSearch!MBMCInvoiceNo)
                If PrevInvNO = "" Then
                    PrevInvNO = Trim(rsSearch!MBMCInvoiceNo)
                ElseIf CurrInvNo <> PrevInvNO Then
                    MultiInvFound = True
                End If
            End If
            
            tmpTotPrinc = tmpTotPrinc + 1
            lblTotPrinc = tmpTotPrinc
            If Trim(rsSearch!MBMCInvoiced) = "Y" Then
                ListMaster.SubItems(10) = "YES"
            Else
                ListMaster.SubItems(10) = "NO"
            End If
            If Trim(rsSearch!INSCode) = "F" Or Trim(rsSearch!INSCode) = "H" Then
                suppExist = False
                tmpSuppID = Trim(rsSearch!MBMNumber)

                If UBound(SuppIDList) > 0 Then
                    For AA = LBound(SuppIDList) To UBound(SuppIDList)
                         If tmpSuppID = SuppIDList(AA) Then
                            suppExist = True
                            Exit For
                         End If
                         
                     Next AA
                End If
                If suppExist = False Then
                    ReDim Preserve SuppIDList(AA + 1)
                    SuppIDList(AA + 1) = tmpSuppID
                End If
                Call InvSuppList(True, rsSearch!MBMNumber)
            End If
        rsSearch.MoveNext
    Loop
    rsSearch.Close
    Set rsSearch = Nothing
    Call InvSuppList(False, "")
    If lvwInvoice.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
    Else
        cmdINVPrint.Enabled = True
        cmdINVPrintPreview.Enabled = True
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboINVCriteriaSel_Change()
    Call INVCriterialSelDisp
End Sub

Private Sub cboINVCriteriaSel_Click()
    Call INVCriterialSelDisp
End Sub

Private Sub cboINVCriteriaSel_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINVCriteriaSel, KeyAscii)
End Sub

Private Sub cboINVListingVer_Click()
Select Case cboINVListingVer.ListIndex
    Case 0
        lbl34.Caption = "Payor"
        cboINVPayor.Visible = True
        cboINVGrpCom.Visible = False
        Label1.Visible = False
        CboDepartment.Visible = False
    Case 1
        lbl34.Caption = "Group Company"
        cboINVGrpCom.Visible = True
        cboINVPayor.Visible = False
        Label1.Visible = False
        CboDepartment.Visible = False
    Case 2
        lbl34.Caption = "Group Company"
        cboINVGrpCom.Visible = True
        cboINVPayor.Visible = False
        Label1.Visible = True
        CboDepartment.Visible = True
    End Select
End Sub

Private Sub cboINVListingVer_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINVListingVer, KeyAscii)
End Sub


Private Sub cboINVPayor_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINVPayor, KeyAscii)
End Sub

Private Sub cboINVPrintOpt_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINVPrintOpt, KeyAscii)
End Sub

Private Sub cmdINVClear_Click()
    txtINVBatchNo = ""
    txtINVBordxDate = "__/__/____"
    txtINVBordxDateFr = "__/__/____"
    txtINVBordxDateTo = "__/__/____"
    cboINVPrintOpt.ListIndex = 0
    txtINVInvoiceNo = ""
    CboDepartment.Text = ""
    lvwInvoice.ListItems.Clear
    cmdINVPrint.Enabled = False
    cmdINVPrintPreview.Enabled = False
    lblTotPrinc = 0
    lblTotSupp = 0
    MultiInvFound = False
    PrevInvNO = ""
    CurrInvNo = ""
End Sub

Private Sub cmdINVClose_Click()
    Unload Me
End Sub

Private Sub cmdINVPrint_Click()
    Call PrintORPreviewInvoice
End Sub

Private Sub cmdINVPrintPreview_Click()
        Call PrintORPreviewInvoice
End Sub

Private Sub cmdINVSearch_Click()
Const METHOD_NAME = "cmdINVSearch_Click"
On Error GoTo debugs

Dim errorFound As Boolean
errorFound = True

    lblTotPrinc = ""
    lblTotSupp = ""
    
    If cboINVCriteriaSel.ListIndex = 0 Then
        If txtINVBordxDate = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Date!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf txtINVBatchNo = "" Then
            MsgBox "Please Enter Batch Number!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    ElseIf cboINVCriteriaSel.ListIndex = 1 Then
        If txtINVBordxDateFr = "__/__/____" Or txtINVBordxDateTo = "__/__/____" Then
            MsgBox "Please enter valid Bordereaux Range !", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    ElseIf cboINVCriteriaSel.ListIndex = 2 Then
        If txtINVInvoiceNo = "" Then
            MsgBox "Please enter Invoice Number!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            errorFound = False
        End If
    End If
    If errorFound = False And cboINVCriteriaSel.ListIndex <> 2 Then
        errorFound = True
        If cboINVListingVer.ListIndex = 0 Then
            If cboINVPayor = "" Then
                MsgBox "Please Select a Payor!", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                errorFound = False
            End If
        ElseIf cboINVListingVer.ListIndex = 1 Or cboINVListingVer.ListIndex = 2 Then
            If cboINVGrpCom = "" Then
                MsgBox "Please Select a Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                errorFound = False
            End If
            If errorFound = False Then
                errorFound = True
                If cboINVListingVer.ListIndex = 2 And CboDepartment = "" Then
                    MsgBox "Please Select a Department!", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    errorFound = False
                End If
            End If
        End If
    End If
    If errorFound = False Then
        tmpTotSupp = 0
        lblTotPrinc = 0
        cmdINVClose.Enabled = False
        MultiInvFound = False
        PrevInvNO = ""
        CurrInvNo = ""
        RecordSearching_INV lvwInvoice
        cmdINVClose.Enabled = True
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If


End Sub

Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

Dim PAY As New ADODB.Recordset
Dim sqlstr As String
Dim GrpCom As New ADODB.Recordset
Dim sqlstr2 As String
   
    sqlstr = "select PAYCode, PAYCompanyName from PAY "
    
    PAY.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    Do While Not PAY.EOF
        cboINVPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
    
    sqlstr = "select GRPCoCode , GRPCoName from GRPCompany"
    
    GrpCom.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    Do While Not GrpCom.EOF
    
        cboINVGrpCom.AddItem Trim((GrpCom!GRPCoCode)) & " - " & Trim((GrpCom!GrpCoName))
        GrpCom.MoveNext
    
    Loop
    
    GrpCom.Close
    Set GrpCom = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmMCOInvoiceNew")
    Call ComboListing
    ReDim SuppIDList(0)
End Sub

Private Sub lvwInvoice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwInvoice.SortKey = ColumnHeader.Index - 1
    lvwInvoice.Sorted = True
    If lvwInvoice.SortOrder = lvwAscending Then
       lvwInvoice.SortOrder = lvwDescending
    Else
       lvwInvoice.SortOrder = lvwAscending
    End If
End Sub

Private Sub UpdateINVPrintingFile(tmpMBMNo As String, tmpCovID As String)
Const METHOD_NAME = "UpdateINVPrintingFile"
On Error GoTo debugs

Dim sqlUpd As String
Dim InvoiceDate As String

    InvoiceDate = Format(Date, "YYYY/MM/DD")
    If tmpCovID = "00" Then
        sqlUpd = " update MBMCLNOthers set " & _
            " MBMCInvoiceDate = '" & InvoiceDate & "', " & _
            " MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "', " & _
            " MBMCInvoiceBy = '" & Logon & "', MBMCInvoiced = 'Y' " & _
            " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
            medixmasterconn.Execute sqlUpd
    Else
        sqlUpd = " update MBMCoveredPersonsCLN set " & _
            " MBMCInvoiceDate = '" & InvoiceDate & "', " & _
            " MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "', " & _
            " MBMCInvoiceBy = '" & Logon & "' " & _
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
Dim strsql As String
Dim InvoiceNo As String
    
    strsql = "Select * from CLNMCOINVQ "
    MBMMCOINVQ.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
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

Private Sub GetSuppMCOFee(PAYCode As String, GRPCoCode As String, MBMDepartment As String, _
MBMBordxDate As String, MBMBatchNo As String, tINSCode As String, TotalPrin, MBMCLNMCOFee)

Const METHOD_NAME = "GetSuppMCOFee"
On Error GoTo debugs

Dim tmpMCO As Double
Dim tmpPrinc As String
Dim strSelFields As String
Dim strSortlist As String
Dim SuppMCORec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim strsql3 As String
Dim strListingV As String
Dim StrReprint As String
Dim strCriterial As String

    strCriterial = ""
    strsql3 = ""
    strListingV = ""
    StrReprint = ""
    If cboINVCriteriaSel.ListIndex = 2 Then
        strCriterial = "AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
    ElseIf cboINVCriteriaSel.ListIndex = 0 Then
        strCriterial = " AND MBMCBordxDate = '" & Format(MBMBordxDate, "mm/dd/yyyy") & "'" & _
                       " AND MBMCBatchNo = " & Trim(MBMBatchNo) & ""
    ElseIf cboINVCriteriaSel.ListIndex = 1 Then
        strCriterial = " AND MBMCBordxDate >= '" & Format(MBMBordxDate, "mm/dd/yyyy") & "'" & _
                  " AND MBMCBordxDate <= '" & Format(MBMBordxDate, "mm/dd/yyyy") & "'"
    End If
    
    If cboINVCriteriaSel.ListIndex <> 2 Then
        If cboINVListingVer = "Invoice - by Payor" Then
            strListingV = " AND PAYCode = '" & Trim(PAYCode) & "'"
        ElseIf cboINVListingVer = "Invoice - by Group Company" Then
            strListingV = " AND GrpCoCode = '" & Trim(GRPCoCode) & "'"
        ElseIf cboINVListingVer = "Invoice - by Department" Then
            strListingV = " AND MBMDepartment = '" & Trim(MBMDepartment) & "'" & _
                          " AND GrpCoCode = '" & Trim(GRPCoCode) & "'"
        End If
    End If
    If cboINVPrintOpt.ListIndex = 0 Then
        StrReprint = " AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
    ElseIf cboINVPrintOpt.ListIndex = 1 Then
        StrReprint = " AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo <> '')"
    End If
    
    strsql3 = strCriterial & " AND INSCode = '" & Trim(tINSCode) & "'" & strListingV & StrReprint
    strSelFields = "SELECT COUNT(MBMCCoverID) AS CQTY, INSCode, SUM(MBMCCLNMCO) AS CMCO," & _
            "MBMCAdultChild, MBMCBordxDate, MBMCBatchNo, MBMCInvoiceNo, MBMDepartment, " & _
            "MBMGroupCompany, MBMCCLNMCO "
    strSortlist = " GROUP BY INSCode, MBMCAdultChild, MBMCBordxDate, MBMCBatchNo, " & _
                "MBMCInvoiceNo , MBMDepartment, MBMGroupCompany, MBMCCLNMCO "
    GetSubSupp = 0
    GetSuppMCO = 0
    tmpMCO = MBMCLNMCOFee
    tmpPrinc = CInt(TotalPrin)
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "MCOFeeNew"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, strSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, strsql3)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelSortList", adVarChar, adParamInput, 2000, strSortlist)
    cmd.Parameters.Append Prm3
    Set Prm4 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 5, 2)
    cmd.Parameters.Append Prm4

    SuppMCORec.CursorLocation = adUseClient
    SuppMCORec.CursorType = adOpenForwardOnly
    SuppMCORec.CacheSize = 100
    Set SuppMCORec = cmd.Execute
    

    tmpPrinc = CInt(TotalPrin)
    Do While Not SuppMCORec.EOF
        tmpPrinc = tmpPrinc + IIf(IsNumeric(SuppMCORec!CQty) = True, SuppMCORec!CQty, 0)
        tmpMCO = tmpMCO + IIf(IsNumeric(SuppMCORec!CMCO) = True, SuppMCORec!CMCO, 0)
        SuppMCORec.MoveNext
    Loop
    SuppMCORec.Close
    Set SuppMCORec = Nothing
    GetSubSupp = tmpPrinc
    GetSuppMCO = tmpMCO
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub PrintORPreviewInvoice()

Const METHOD_NAME = "PrintORPreviewInvoice"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim Cnt As Integer
Dim TotalPrin As String
Dim MemberNo As String
Dim CovID  As String
Dim InsDesc As String
Dim CompTrans As Boolean
Dim TransStart As Boolean
Dim tmpDept As String
Dim tmpPayor As String
Dim tmpCompName As String
Dim tmpAdd1 As String
Dim tmpAdd2 As String
Dim tmpAdd3 As String
Dim tmpContact As String
Dim tmpTelNo As String
Dim tmpFaxNo As String
Dim PrevKey As String
Dim CurrKey As String
Dim MCORec As New ADODB.Recordset
Dim TotCovered As Integer
Dim TOtMco As Double
Dim RecNotFound As Boolean
    If MultiInvFound = True Then
        MsgBox "Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! ", vbCritical + vbOKOnly
        Exit Sub
    End If
    TotalPrin = 0
    Screen.MousePointer = vbHourglass
    TransStart = False
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        Exit Sub
    End If
    
    On Error GoTo HANDLE_ERROR
    CompTrans = False
    objExcel.Interactive = False
    Set objExcel = New Excel.Application
    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    RecNotFound = False
    Call ReadRecord(MCORec, 3, "P", RecNotFound, False, False, "")
    If RecNotFound = True Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboINVListingVer = "Invoice - by Payor" Then
        tmpPayor = IIf(Len(MCORec!PAYCode), MCORec!PAYCode, "")
    Else
        tmpPayor = IIf(Len(MCORec!GRPCoCode), MCORec!GRPCoCode, "")
    End If

    objExcel.WindowState = xlMaximized

    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MCOInvoice\MCOInvoice.xlt")

    Set objWorksheet = objWorkbook.Sheets(1)
    objExcel.DisplayAlerts = False
        
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    
    Call PrintPAYInfo(tmpPayor, tmpCompName, tmpAdd1, tmpAdd2, tmpAdd3, tmpContact, tmpTelNo, tmpFaxNo)
    objWorksheet.Cells(8, "A") = tmpCompName
    objWorksheet.Cells(9, "A") = tmpAdd1
    objWorksheet.Cells(10, "A") = tmpAdd2
    objWorksheet.Cells(11, "A") = tmpAdd3
    objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & tmpContact
    objWorksheet.Cells(15, "A") = "TEL :" & " " & tmpTelNo
    objWorksheet.Cells(15, "C") = "FAX :" & " " & tmpFaxNo
    tmpDept = ""
    If cboINVListingVer = "Invoice - by Department" Then
        tmpDept = IIf(Len(MCORec!MBMDepartment), Trim(MCORec!MBMDepartment), "")
        objWorksheet.Cells(7, "A") = tmpDept
    End If
            
    If Trim(MCORec!MBMCInvoiceNo) <> "" Then
        objWorksheet.Cells(8, "G") = Trim(MCORec!MBMCInvoiceNo)
    End If
    If Len(MCORec!MBMCInvoiceDate) Then
        objWorksheet.Cells(10, "G") = MCORec!MBMCInvoiceDate
    Else
        objWorksheet.Cells(10, "G") = Format(Date, "dd/mm/yyyy")
    End If
    Cnt = 0
    Do Until MCORec.EOF
        If Cnt = 0 Then
            PrevKey = Trim(MCORec!INSCode) & "*" & Format(MCORec!MBMBordxDate, "ddmmyyyy") & "*" & Trim(MCORec!MBMAdultChild)
        Else
            CurrKey = Trim(MCORec!INSCode) & "*" & Format(MCORec!MBMBordxDate, "ddmmyyyy") & "*" & Trim(MCORec!MBMAdultChild)
        End If
        If PrevKey <> CurrKey Then
            TotalPrin = 0
            If Trim(MCORec!MBMMemberType) <> "S" Then
                TotalPrin = MCORec!MBMQty
            End If
            TotCovered = MCORec!MBMQty
            TOtMco = MCORec!MCOFee
            PrevKey = Trim(MCORec!INSCode) & "*" & Format(MCORec!MBMBordxDate, "ddmmyyyy") & "*" & Trim(MCORec!MBMAdultChild)
            Cnt = Cnt + 1
        Else
            TotCovered = TotCovered + MCORec!MBMQty
            TOtMco = TOtMco + MCORec!MCOFee
        End If

        If Trim(MCORec!INSCode) = "I" Then
            InsDesc = "I - INDIVIDUAL"
        ElseIf Trim(MCORec!INSCode) = "F" Then
            InsDesc = "F - FAMILY"
        ElseIf Trim(MCORec!INSCode) = "G" Then
            InsDesc = "G - GROUP INDIVIDUAL"
        ElseIf Trim(MCORec!INSCode) = "H" Then
            InsDesc = "H - GROUP FAMILY"
        End If
        objWorksheet.Cells(20 + CDbl(Cnt), "A") = Cnt
        objWorksheet.Cells(20 + CDbl(Cnt), "B") = Format(MCORec!MBMBordxDate, "mm/dd/yyyy")
        objWorksheet.Cells(20 + CDbl(Cnt), "C") = InsDesc
        
        If Trim(MCORec!MBMAdultChild) = "A" Then
            objWorksheet.Cells(20 + CDbl(Cnt), "D") = "Adult"
        ElseIf Trim(MCORec!MBMAdultChild) = "C" Then
            objWorksheet.Cells(20 + CDbl(Cnt), "D") = "Child"
        End If
        objWorksheet.Cells(20 + CDbl(Cnt), "E") = TotalPrin
        objWorksheet.Cells(20 + CDbl(Cnt), "F") = TotCovered
        objWorksheet.Cells(20 + CDbl(Cnt), "G") = TOtMco
        MCORec.MoveNext
    Loop
    MCORec.Close
    Set MCORec = Nothing
            
    objWorksheet.Range("E43").Formula = "=SUM(E21:E41)"
    objWorksheet.Range("F43").Formula = "=SUM(F21:F41)"
    objWorksheet.Range("G43").Formula = "=SUM(G21:G41)"
    
    
    If cmdINVPrint = True Then
        If Trim(txtInvoiceNo) = "" Then
            TransStart = True
            MediConnCISDB.beginTrans
            medixmasterconn.beginTrans
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                Call IncInvoiceNo
                objWorksheet.Cells(8, "G") = Trim(txtInvoiceNo)
            ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                Call IncInvoiceNo
                objWorksheet.Cells(8, "G") = Trim(txtInvoiceNo)
            End If
            For Cnt = 1 To lvwInvoice.ListItems.Count
                If Trim(lvwInvoice.ListItems(Cnt).SubItems(10)) = "NO" Then
                    lvwInvoice.ListItems(Cnt).SubItems(10) = "YES"
                    lvwInvoice.ListItems(Cnt).SubItems(11) = txtInvoiceNo
                    MemberNo = Trim(lvwInvoice.ListItems(Cnt).SubItems(1))
                    CovID = Trim(lvwInvoice.ListItems(Cnt).SubItems(2))
                    Call UpdateINVPrintingFile(MemberNo, CovID)
                End If
            Next Cnt
            medixmasterconn.commitTrans
            MediConnCISDB.commitTrans
            CompTrans = True
            cmdINVPrint.Enabled = False
            cboINVPrintOpt.ListIndex = 1
        End If
    End If
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Screen.MousePointer = vbDefault
    Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            If CompTrans = False And TransStart = True Then
                medixmasterconn.RollbackTrans
                MediConnCISDB.RollbackTrans
            End If
            Screen.MousePointer = vbDefault
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
            
End Sub

Private Sub INVCriterialSelDisp()
Const METHOD_NAME = "INVCriterialSelDisp"
On Error GoTo debugs

    Call cmdINVClear_Click
    txtINVBatchNo.Visible = False
    txtINVBordxDateFr.Visible = False
    txtINVBordxDateTo.Visible = False
    txtINVInvoiceNo.Visible = False
    cboINVPayor.Enabled = True
    cboINVGrpCom.Enabled = True
    txtINVBordxDate.Visible = False
    txtINVBordxDateFr.Visible = False
    txtINVBordxDateTo.Visible = False
    cboINVPrintOpt.Enabled = False

    Select Case cboINVCriteriaSel.ListIndex
        Case 0
            lbl29.Caption = "Bordx Date"
            lbl30.Caption = "Batch Num"
            txtINVBordxDate.Visible = True
            txtINVBatchNo.Visible = True
            cboINVPrintOpt.ListIndex = 0
            cboINVPrintOpt.Enabled = True
            cboINVListingVer.Enabled = True
            cboINVListingVer.ListIndex = 0
        Case 1
            lbl29.Caption = "Bordx From"
            lbl30.Caption = "To"
            txtINVBordxDateFr.Visible = True
            txtINVBordxDateTo.Visible = True
            cboINVPrintOpt.ListIndex = 0
            cboINVPrintOpt.Enabled = True
            cboINVListingVer.Enabled = True
            cboINVListingVer.ListIndex = 0
        Case 2
            lbl29.Caption = "Invoice Num"
            lbl30.Caption = ""
            txtINVInvoiceNo.Visible = True
            cboINVPrintOpt.ListIndex = 1
            cboINVListingVer.ListIndex = 1
            cboINVPayor.Enabled = False
            cboINVGrpCom.Enabled = False
            lvwInvoice.ListItems.Clear
            lblTotPrinc = 0
            lblTotSupp = 0
    End Select

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub InvSuppList(tmpFrMBMNo As Boolean, tmpMBMNo As String)
Const METHOD_NAME = "InvSuppList"
On Error GoTo debugs

Dim MBMCov As New ADODB.Recordset
Dim ListMaster As ListItem
Dim tmpMBMID As String
Dim AA As Integer
Dim suppExist As Boolean
    
    Set MBMCov = Nothing
    Call ReadRecord(MBMCov, 2, "S", False, True, tmpFrMBMNo, tmpMBMNo)
    Do Until MBMCov.EOF
        suppExist = False
        If tmpFrMBMNo = False Then
            tmpMBMID = Trim(MBMCov!MBMCNUmber)
            If UBound(SuppIDList) > 0 Then
                For AA = LBound(SuppIDList) To UBound(SuppIDList)
                     If tmpMBMID = Trim(SuppIDList(AA)) Then
                        suppExist = True
                        Exit For
                     End If
                 Next AA
            End If
        End If
        'MBMGroupCompany, MBMCBordxDate, MBMCBatchNo, MBMCInvoiceNo "
        If suppExist = False Then
           Set ListMaster = lvwInvoice.ListItems.Add(, , MBMCov!PAYCode)
            ListMaster.SubItems(1) = Trim(MBMCov!MBMCNUmber)
            ListMaster.SubItems(2) = MBMCov!MBMCCoverID
            ListMaster.SubItems(3) = MBMCov!MBMCRELCode
            ListMaster.SubItems(4) = MBMCov!MBMCBordxDate
            ListMaster.SubItems(5) = IIf(Len(MBMCov!MBMPolicyNo), Trim(MBMCov!MBMPolicyNo), "")
            ListMaster.SubItems(7) = IIf(Len(MBMCov!MBMCBatchNo), MBMCov!MBMCBatchNo, "")
            ListMaster.SubItems(8) = Trim(MBMCov!MBMCName)
           If MBMCov!MBMGroupCompany = "" Then
                ListMaster.SubItems(9) = ""
            Else
                If Len(MBMCov!MBMGroupCompany) Then
                    ListMaster.SubItems(9) = Trim(MBMCov!MBMGroupCompany)
                End If
            End If
        
           If Len(MBMCov!MBMCInvoiceNo) Then
               txtInvoiceNo = Trim(MBMCov!MBMCInvoiceNo)
               ListMaster.SubItems(11) = Trim(MBMCov!MBMCInvoiceNo)
               ListMaster.SubItems(10) = "YES"
               CurrInvNo = Trim(MBMCov!MBMCInvoiceNo)
               If PrevInvNO = "" Then
                    PrevInvNO = Trim(MBMCov!MBMCInvoiceNo)
               ElseIf CurrInvNo <> PrevInvNO Then
                    MultiInvFound = True
               End If
           Else
               ListMaster.SubItems(10) = "NO"
           End If
                    
           tmpTotSupp = tmpTotSupp + 1
           lblTotSupp = tmpTotSupp
           
        End If
        MBMCov.MoveNext
    Loop
    MBMCov.Close
    Set MBMCov = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function SelectionCriterial(tmpMBMType As String) As String

Const METHOD_NAME = "SelectionCriterial"
On Error GoTo debugs

Dim strListingV As String
Dim StrReprint As String
Dim strCriterial As String
Dim tmpInvoice As String
Dim tmpBordxDate As String
Dim tmpBatchNo As String

Dim formattmpBordxDate As String
Dim formatINVBordxDateTo As String
Dim formatINVBordxDateFr As String

    If tmpMBMType = "P" Then
        tmpBordxDate = "MBMBordxDate"
        tmpBatchNo = "MBMBatchNo"
    Else
        tmpBordxDate = "MBMCBordxDate"
        tmpBatchNo = "MBMCBatchNo"
    End If
    If cboINVCriteriaSel.ListIndex = 2 Then
        strCriterial = " AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
    ElseIf cboINVCriteriaSel.ListIndex = 0 Then
        'shan 08092009
        'formattmpBordxDate = "Convert(VarChar(10), " & txtINVBordxDate & ", 101)"
        'strCriterial = " AND (" & tmpBordxDate & " = '" & Format(formattmpBordxDate, "mm/dd/yyyy") & "'" & _
        '           " AND " & tmpBatchNo & " = " & Trim(txtINVBatchNo) & ")"
        strCriterial = " AND (" & tmpBordxDate & " = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                   " AND " & tmpBatchNo & " = " & Trim(txtINVBatchNo) & ")"
    ElseIf cboINVCriteriaSel.ListIndex = 1 Then
        'shan 08092009
        'formatINVBordxDateFr = "Convert(VarChar(10), " & txtINVBordxDateFr & ", 101)"
        'formatINVBordxDateTo = "Convert(VarChar(10), " & txtINVBordxDateTo & ", 101)"
        
        'strCriterial = " AND " & tmpBordxDate & " >= '" & Format(formatINVBordxDateFr, "mm/dd/yyyy") & "'" & _
              " AND " & tmpBordxDate & " <= '" & Format(formatINVBordxDateTo, "mm/dd/yyyy") & "'"
        strCriterial = " AND " & tmpBordxDate & " >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
              " AND " & tmpBordxDate & " <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'"
    End If
    If cboINVCriteriaSel.ListIndex <> 2 Then
        If cboINVListingVer = "Invoice - by Payor" Then
            strListingV = " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
        ElseIf cboINVListingVer = "Invoice - by Group Company" Then
            strListingV = " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
        ElseIf cboINVListingVer = "Invoice - by Department" Then
            strListingV = " AND MBMDepartment = '" & Trim(CboDepartment) & "'" & _
                          " AND GrpCoCode = '" & Mid(cboINVGrpCom, 1, 4) & "'"
        End If
    End If
    If cboINVPrintOpt.ListIndex = 0 Then
        StrReprint = " AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
    ElseIf cboINVPrintOpt.ListIndex = 1 Then
        StrReprint = " AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo <> '')"
    End If
    SelectionCriterial = strCriterial & strListingV & StrReprint
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub ReadRecord(rsSearch As ADODB.Recordset, strAppendCase As String, DataType As String, RecNotFound As Boolean, _
ListOPt As Boolean, tmpFrMBMNo As Boolean, tmpExtraCri As String)

Const METHOD_NAME = "ReadRecord"
On Error GoTo debugs


Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim strSelFields As String
Dim strSortlist As String
Dim tmpSelCri As String
    
    tmpSelCri = SelectionCriterial(DataType)
    strSortlist = ""
    If ListOPt = True Then
        If DataType = "P" Then
            strSelFields = "SELECT PAYCode, MBMNumber, MBMMemberType, MBMBordxDate, MBMPolicyNo, " & _
                           "MBMBatchNo, MBMName, MBMGroupCompany, MBMCInvoiceNo, MBMCInvoiced, INSCode"
            strSortlist = ""
            tmpSelCri = tmpSelCri & " ORDER BY MBMNumber"
        Else
            strSelFields = "SELECT PAYCode, MBMCNumber,MBMCCoverID,MBMCRELCode, MBMCName,MBMGroupCompany, " & _
                           "MBMCBordxDate, MBMCBatchNo, MBMCInvoiceNo, MBMPolicyNo "
            If tmpFrMBMNo = True Then
                tmpSelCri = tmpSelCri & " AND MBMCNumber='" & tmpExtraCri & "' "
            End If
            tmpSelCri = tmpSelCri & " ORDER BY MBMCNumber "
        End If
    Else
        strSelFields = "SELECT INSCode, MBMBordxDate, MBMMemberType,MBMCInvoiceNo, MBMCInvoiceDate, " & _
                       "PAYCode, GRPCoCode, MBMAdultChild, SUM(MBMQty) AS MBMQty, SUM(MCOFee) AS MCOFee"
    
        strSortlist = "GROUP BY INSCode, MBMBordxDate, MBMMemberType, MBMCInvoiceNo , MBMCInvoiceDate, PAYCode, GRPCoCode, MBMAdultChild"
    End If
    If cboINVListingVer = "Invoice - by Department" Then
        strSelFields = strSelFields & ", MBMDepartment "
        strSortlist = strSortlist & ", MBMDepartment "
    End If
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "MCOFeeNew"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, strSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, tmpSelCri)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelSortList", adVarChar, adParamInput, 2000, strSortlist)
    cmd.Parameters.Append Prm3
    Set Prm4 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 5, strAppendCase)
    cmd.Parameters.Append Prm4

    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    rsSearch.CacheSize = 100
    Set rsSearch = cmd.Execute
    RecNotFound = False
    If rsSearch.EOF Then RecNotFound = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub PrintPAYInfo(tmpPayor As String, tmpCompName As String, tmpAdd1 As String, tmpAdd2 As String, _
tmpAdd3 As String, tmpContact As String, tmpTelNo As String, tmpFaxNo As String)

Const METHOD_NAME = "PrintPAYInfo"
On Error GoTo debugs

Dim tmpPayGrp As New ADODB.Recordset
Dim strSqlPay As String

    If cboINVListingVer = "Invoice - by Payor" Then
        strSqlPay = "SELECT PAYCompanyName AS CompanyName,PAYAddress1 as Add1, " & _
                    "PAYAddress2 as Add2,PAYAddress3 as Add3, PAYMAINContact as Contact, " & _
                    " PAYTELno as TelNo, PAYFAXNo as FaxNo " & _
                    " From PAY " & _
                    " WHERE PAYCODE = '" & tmpPayor & "'"
    Else
        strSqlPay = "SELECT GRPCoName AS CompanyName,GRPCoAdd1 as Add1, " & _
                    "GRPCoAdd2 as Add2,GRPCoAdd3 as Add3, " & _
                    "GRPCoContact as Contact, GRPCoPhone as TelNo, GRPCoFax as FaxNo " & _
                    " From GRPCompany " & _
                    " WHERE GRPCoCode = '" & tmpPayor & "'"
    End If
    tmpPayGrp.Open strSqlPay, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
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

