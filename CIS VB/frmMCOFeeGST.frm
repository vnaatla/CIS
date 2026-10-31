VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCOFeeGST 
   Caption         =   "Form1"
   ClientHeight    =   8064
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11376
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8064
   ScaleWidth      =   11376
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbosubsidiary 
      Height          =   315
      Left            =   6000
      TabIndex        =   33
      Top             =   1800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame FrameINV 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   11055
      Begin VB.TextBox txtINVBatchNo 
         Height          =   315
         Left            =   8520
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtINVInvoiceNo 
         Height          =   315
         Left            =   6000
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboINVCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCOFeeGST.frx":0000
         Left            =   2160
         List            =   "frmMCOFeeGST.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cboINVListingVer 
         Height          =   315
         ItemData        =   "frmMCOFeeGST.frx":0041
         Left            =   2160
         List            =   "frmMCOFeeGST.frx":004B
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   650
         Width           =   5415
      End
      Begin VB.ComboBox cboINVPayor 
         Height          =   315
         Left            =   2160
         TabIndex        =   17
         Text            =   "cboINVPayor"
         Top             =   930
         Width           =   5415
      End
      Begin VB.ComboBox cboINVGrpCom 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   930
         Width           =   5415
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         Left            =   6000
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cboINVPrintOpt 
         Height          =   315
         ItemData        =   "frmMCOFeeGST.frx":007F
         Left            =   2160
         List            =   "frmMCOFeeGST.frx":0089
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1200
         Width           =   2415
      End
      Begin MSMask.MaskEdBox txtINVBordxDate 
         Height          =   315
         Left            =   6000
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   24
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
      Begin VB.Label lblSubsidiary 
         Caption         =   "SubSidiary"
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl29 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   4920
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   1260
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl30 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   7560
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl36 
         Caption         =   "Print Option"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl35 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lbl34 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lbl27 
         Caption         =   "Single/Batch/Range"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
   End
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
         Left            =   480
         TabIndex        =   12
         Top             =   0
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   11055
      Begin VB.TextBox lblTotSupp 
         Height          =   285
         Left            =   3840
         TabIndex        =   40
         Text            =   "0"
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox lblTotPrinc 
         Height          =   285
         Left            =   1440
         TabIndex        =   39
         Text            =   "0"
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton cmdINVClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9240
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrintPreview 
         Caption         =   "Pre&view"
         Enabled         =   0   'False
         Height          =   300
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7560
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   300
         Left            =   6720
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   285
         Left            =   5880
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdINVClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   3
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   240
         Width           =   1200
      End
   End
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
      TabIndex        =   0
      Top             =   2400
      Width           =   11055
      Begin MSComctlLib.ListView lvwInvoice 
         Height          =   4635
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10815
         _ExtentX        =   19071
         _ExtentY        =   8170
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
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Subsidiary"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Print Option"
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Print Option"
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Print Option"
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Print Option"
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Department"
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmMCOFeeGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SuppIDList() As String
Dim tmpTotSupp As Long
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
    Call BindSubsidary
    
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

Public Sub BindSubsidary()
  Dim rst3 As New ADODB.Recordset
  Dim sql2 As String
  
  Dim grpComp() As String
  grpComp = Split(cboINVGrpCom, "-")
  
  
  
        sql2 = ""
        sql2 = "Exec SearchSubsidiary '" & Trim(Replace(grpComp(1), "'", "''")) & "'"
        rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
  
        cbosubsidiary.Clear
        Do Until rst3.EOF
            With rst3
                cbosubsidiary.AddItem IIf(Len(Trim(!MBMGRPSubsidiary)), Trim(!MBMGRPSubsidiary), "")
                rst3.MoveNext
            End With
        Loop
        rst3.Close
        Set rst3 = Nothing

End Sub

'Private Sub cboINVPayor_LostFocus()
   ' Call BindSubsidary
'End Sub

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
'Private Sub RecordSearching_INV(objListView As Object)
'    Screen.MousePointer = vbHourglass
'        Call INVPrincipalListing
'    Screen.MousePointer = Default
'End Sub

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
    Call ReadRecord(rsSearch)
        Do Until rsSearch.EOF
          With rsSearch
            Set ListMaster = lvwInvoice.ListItems.Add(, , IIf(Len(!PAYCode), !PAYCode, ""))
            ListMaster.SubItems(1) = Trim(!MBMCNUmber)
            ListMaster.SubItems(2) = Trim(!MBMCCoverID)
            ListMaster.SubItems(3) = Trim(!RelCode)
            ListMaster.SubItems(4) = Trim(!MBMBordxDate)
            If Len(!MBMPolicyNo) Then
                ListMaster.SubItems(5) = Trim(!MBMPolicyNo)
            End If
            If Len(!MBMBatchNo) Then
                ListMaster.SubItems(7) = !MBMBatchNo
            End If
            ListMaster.SubItems(8) = Trim(!MBMName)
            If Len(Trim(!MBMGroupCompany)) = 0 Or IsNull(!MBMGroupCompany) = True Then
                ListMaster.SubItems(9) = ""
            Else
                ListMaster.SubItems(9) = Trim(!MBMGroupCompany)
            End If
            If Len(!MBMCInvoiceNo) Then
                txtInvoiceNo = Trim(!MBMCInvoiceNo)
                ListMaster.SubItems(11) = Trim(!MBMCInvoiceNo)
                CurrInvNo = Trim(!MBMCInvoiceNo)
                If PrevInvNO = "" Then
                    PrevInvNO = Trim(!MBMCInvoiceNo)
                ElseIf CurrInvNo <> PrevInvNO Then
                    MultiInvFound = True
                End If
            End If
             If Len(Trim(cbosubsidiary)) Then
                If Len(!Subsidiary) Then
                    ListMaster.SubItems(12) = IIf(Len(Trim(!Subsidiary)), Trim(!Subsidiary), "")
                End If
            Else
                ListMaster.SubItems(12) = ""
                End If
            If Trim(!RelCode) = Trim("P") Then
                tmpTotPrinc = tmpTotPrinc + 1
                lblTotPrinc = tmpTotPrinc
            Else
               tmpTotSupp = tmpTotSupp + 1
               lblTotSupp = tmpTotSupp
            End If
            If Len(!MBMCInvoiceNo) Then
                ListMaster.SubItems(10) = "YES"
            Else
                ListMaster.SubItems(10) = "NO"
            End If
            
            
        rsSearch.MoveNext
        End With
    Loop
    
    rsSearch.Close
    Set rsSearch = Nothing
   ' Call InvSuppList(False, "")
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
         lblSubsidiary.Visible = False
        cbosubsidiary.Visible = False
    Case 1
        lbl34.Caption = "Group Company"
        cboINVGrpCom.Visible = True
        cboINVPayor.Visible = False
        Label1.Visible = False
        CboDepartment.Visible = False
        lblSubsidiary.Visible = True
        cbosubsidiary.Visible = True
    Case 2
        lbl34.Caption = "Group Company"
        cboINVGrpCom.Visible = True
        cboINVPayor.Visible = False
        Label1.Visible = True
        CboDepartment.Visible = True
         lblSubsidiary.Visible = False
        cbosubsidiary.Visible = False
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
Dim Cnt As Integer
Dim MemberNo As String
Dim CovID As String

        If Trim(txtInvoiceNo) = "" Then
            If cboINVCriteriaSel.ListIndex = 0 And cboINVPrintOpt.ListIndex = 0 Then
                Call IncInvoiceNo
            ElseIf cboINVCriteriaSel.ListIndex = 1 Then
                Call IncInvoiceNo
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
            cmdINVPrint.Enabled = False
            cboINVPrintOpt.ListIndex = 1
        End If
        Call PrintInvoice(txtInvoiceNo)
End Sub

Public Sub PrintInvoice(InvoiceNo As String)
Const METHOD_NAME = "PrintInvoice"
On Error GoTo debugs

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim tmpbordxdate As String
Dim tmpBatchNo As String
Dim sqlString As String
sqlString = ""
sqlString = " AND MBMCInvoiceNo = '" & Trim(InvoiceNo) & "'"

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim rsSearch As New ADODB.Recordset
Dim strSelFields As String
Dim strSortlist As String
Dim tmpSelCri As String
Dim DataType As String
Dim IsFrst As Boolean
Dim tmpPayor As String
Dim InsDesc As String
Dim Cnt As Integer
IsFrst = True

    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SP_MCOSummaryGSTOP"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, sqlString)
    cmd.Parameters.Append Prm1
    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    rsSearch.CacheSize = 100
    Set rsSearch = cmd.Execute
    Cnt = 1
    Do Until rsSearch.EOF
    
        If IsFrst Then
            If objExcel Is Nothing Then
                Set objExcel = CreateObject("Excel.application")
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MCOInvoice\MCOInvoice.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
            tmpPayor = IIf(Len(rsSearch!PAYCode), rsSearch!PAYCode, "")
            Dim tmpCompName As String
            Dim tmpAdd1 As String
            Dim tmpAdd2 As String
            Dim tmpAdd3 As String
            Dim tmpContact As String
            Dim tmpTelNo As String
            Dim tmpFaxNo As String
            Call PrintPAYInfo(tmpPayor, tmpCompName, tmpAdd1, tmpAdd1, tmpAdd3, tmpContact, tmpTelNo, tmpFaxNo)
            objWorksheet.Cells(8, "A") = tmpCompName
            objWorksheet.Cells(9, "A") = tmpAdd1
            objWorksheet.Cells(10, "A") = tmpAdd2
            objWorksheet.Cells(11, "A") = tmpAdd3
            objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & tmpContact
            objWorksheet.Cells(15, "A") = "TEL :" & " " & tmpTelNo
            objWorksheet.Cells(15, "C") = "FAX :" & " " & tmpFaxNo
            If Trim(rsSearch!MBMCInvoiceNo) <> "" Then
                objWorksheet.Cells(8, "G") = Trim(rsSearch!MBMCInvoiceNo)
            End If
            If cboINVListingVer = "Invoice - by Payor" Then
                tmpPayor = IIf(Len(rsSearch!PAYCode), rsSearch!PAYCode, "")
            Else
                tmpPayor = IIf(Len(rsSearch!GRPCoCode), rsSearch!GRPCoCode, "")
            End If
            IsFrst = False
        End If
        
        If Trim(rsSearch!INSCode) = "I" Then
            InsDesc = "I - INDIVIDUAL"
        ElseIf Trim(rsSearch!INSCode) = "F" Then
            InsDesc = "F - FAMILY"
        ElseIf Trim(rsSearch!INSCode) = "G" Then
            InsDesc = "G - GROUP INDIVIDUAL"
        ElseIf Trim(rsSearch!INSCode) = "H" Then
            InsDesc = "H - GROUP FAMILY"
        End If
        objWorksheet.Cells(20 + CDbl(Cnt), "A") = Cnt
        objWorksheet.Cells(20 + CDbl(Cnt), "B") = Format(rsSearch!MBMBordxDate, "mm/dd/yyyy")
        objWorksheet.Cells(20 + CDbl(Cnt), "C") = InsDesc
        
        If Trim(rsSearch!AdultChild) = "A" Then
            objWorksheet.Cells(20 + CDbl(Cnt), "D") = "Adult"
        ElseIf Trim(rsSearch!AdultChild) = "C" Then
            objWorksheet.Cells(20 + CDbl(Cnt), "D") = "Child"
        End If
        objWorksheet.Cells(20 + CDbl(Cnt), "E") = rsSearch!Empcnt
        objWorksheet.Cells(20 + CDbl(Cnt), "F") = rsSearch!DepCnt
        objWorksheet.Cells(20 + CDbl(Cnt), "G") = rsSearch!BasicMCO
        objWorksheet.Cells(20 + CDbl(Cnt), "H") = IIf(Len(rsSearch!GST), rsSearch!GST, 0)
        Cnt = Cnt + 1
        rsSearch.MoveNext
    Loop
        
    objWorksheet.Range("E43").Formula = "=SUM(E21:E41)"
    objWorksheet.Range("F43").Formula = "=SUM(F21:F41)"
    objWorksheet.Range("G43").Formula = "=SUM(G21:G41)"
    objWorksheet.Range("H43").Formula = "=SUM(H21:H41)"
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Screen.MousePointer = vbDefault
    
    Dim MCOdetails As New ADODB.Recordset
    Dim Parm1 As New ADODB.Parameter
    Dim McoDetCmd As New ADODB.Command
    Dim Rcnt As Integer
    
    IsFrst = True
    Set McoDetCmd.ActiveConnection = MediConnCISDB
    McoDetCmd.CommandText = "SP_MCOGSTInvoiceOP"
    McoDetCmd.CommandType = adCmdStoredProc
    McoDetCmd.CommandTimeout = 300
    Set Prm1 = McoDetCmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, sqlString)
    McoDetCmd.Parameters.Append Prm1
    MCOdetails.CursorLocation = adUseClient
    MCOdetails.CursorType = adOpenForwardOnly
    MCOdetails.CacheSize = 100
    Set MCOdetails = McoDetCmd.Execute
    Rcnt = 4
    Cnt = 1
    Do Until MCOdetails.EOF
        With MCOdetails
        If IsFrst Then
            Set objWorksheet = objWorkbook.Sheets(2)
            Set objWorksheet = objWorkbook.Worksheets.Item(2)
            objExcel.DisplayAlerts = False
            objWorksheet.Cells(Rcnt, "A") = "No"
            objWorksheet.Cells(Rcnt, "B") = "Payor Code"
            objWorksheet.Cells(Rcnt, "C") = "MBM Number"
            objWorksheet.Cells(Rcnt, "D") = "Cover ID"
            objWorksheet.Cells(Rcnt, "E") = "Relationship"
            objWorksheet.Cells(Rcnt, "F") = "Bordx Date"
            objWorksheet.Cells(Rcnt, "G") = "Policy Number"
            objWorksheet.Cells(Rcnt, "H") = "Health Type"
            objWorksheet.Cells(Rcnt, "I") = "Batch No"
            objWorksheet.Cells(Rcnt, "J") = "Name"
            objWorksheet.Cells(Rcnt, "K") = "Group Company"
            objWorksheet.Cells(Rcnt, "L") = "Invoice Number"
            objWorksheet.Cells(Rcnt, "M") = "Basic MCO"
            objWorksheet.Cells(Rcnt, "N") = "GST"
            Rcnt = Rcnt + 1
            IsFrst = False
        End If
        
            Rcnt = Rcnt + 1
            objWorksheet.Cells(Rcnt, "A") = Cnt
            objWorksheet.Cells(Rcnt, "B") = IIf(Len(!PAYCode), !PAYCode, "")
            objWorksheet.Cells(Rcnt, "C") = IIf(Len(!MBMCNUmber), !MBMCNUmber, "")
            objWorksheet.Cells(Rcnt, "D") = IIf(Len(!MBMCCoverID), !MBMCCoverID, "")
            objWorksheet.Cells(Rcnt, "E") = IIf(Len(!MBMCRELCode), !MBMCRELCode, "")
            objWorksheet.Cells(Rcnt, "F") = IIf(Len(!MBMBordxDate), !MBMBordxDate, "")
            objWorksheet.Cells(Rcnt, "G") = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
            objWorksheet.Cells(Rcnt, "H") = IIf(Len(!HLTCode), !HLTCode, "")
            objWorksheet.Cells(Rcnt, "I") = IIf(Len(!MBMBatchNo), !MBMBatchNo, "")
            objWorksheet.Cells(Rcnt, "J") = IIf(Len(!MBMName), !MBMName, "")
            objWorksheet.Cells(Rcnt, "K") = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
            objWorksheet.Cells(Rcnt, "L") = IIf(Len(!MBMCInvoiceNo), !MBMCInvoiceNo, "")
            objWorksheet.Cells(Rcnt, "M") = IIf(Len(!BasicMCO), !BasicMCO, "")
            objWorksheet.Cells(Rcnt, "N") = IIf(Len(!GST), !GST, "")
            Cnt = Cnt + 1
        End With
       MCOdetails.MoveNext
    Loop
    
         Dim fPDF As String
         Dim DirPath As String
            If Len(Dir("C:\GSTMCOReports\", vbDirectory)) = 0 Then
               MkDir "C:\GSTMCOReports\"
            End If
         DirPath = "C:\GSTMCOReports\"
         If Len(InvoiceNo) Then
            fPDF = DirPath + InvoiceNo + "_" + Replace(CStr(DateTime.Time), ":", "_") + ")" + ".xls"
         Else
            fPDF = DirPath + tmpPayor + "_" + Replace(CStr(DateTime.Time), ":", "_") + ")" + ".xls"
         End If
         objWorksheet.SaveAs fPDF
         objWorkbook.Saved = True
         objWorkbook.Close
         objExcel.Quit
         Set objWorksheet = Nothing
         Set objExcel = Nothing
         MsgBox "Report Generated successfully. Please check on C:\GSTMCOReports " & vbCrLf & vbCrLf
          
    Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            'If CompTrans = False And TransStart = True Then
                'medixmasterconn.RollbackTrans
                'MediConnCISDB.RollbackTrans
            'End If
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
           
End Sub

Private Sub cmdINVPrintPreview_Click()
        Call PrintPreview
End Sub

Private Sub cmdINVSearch_Click()
'Test Trinath New Prog Code
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
        Call INVPrincipalListing
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
    Call ActionLogUpdate(12, "frmMCOFeeGST")
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
            " MBMCInvoiceBy = '" & Logon & "', MBMCInvoiced = 'Y', MBMCGSTAmt= (MBMCLNMCOFee * (0.06)),MBMCGSTPerc='6.0' " & _
            " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
            medixmasterconn.Execute sqlUpd
    Else
        sqlUpd = " update MBMCoveredPersonsCLN set " & _
            " MBMCInvoiceDate = '" & InvoiceDate & "', " & _
            " MBMCInvoiceNo = '" & Trim(txtInvoiceNo) & "', MBMCGSTAmt= (MBMCCLNMCO * (0.06)) , MBMCGSTPerc='6.0' " & _
            " ,MBMCInvoiceBy = '" & Logon & "' " & _
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
Dim prm4 As New ADODB.Parameter
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
    Set prm4 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 5, 2)
    cmd.Parameters.Append prm4

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

Private Sub PrintPreview()

Const METHOD_NAME = "PrintPreview"
On Error GoTo debugs


Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet

Dim tmpbordxdate As String
Dim tmpBatchNo As String
Dim sqlString As String
sqlString = ""

If cboINVCriteriaSel.ListIndex = 0 Then
    sqlString = " AND ( MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
               " AND  MBMBatchNo = " & Trim(txtINVBatchNo) & ")"
ElseIf cboINVCriteriaSel.ListIndex = 1 Then
    sqlString = " AND MBMBordxDate >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
          " AND MBMBordxDate <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'"
ElseIf cboINVCriteriaSel.ListIndex = 2 Then
    sqlString = " AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
End If

If cboINVCriteriaSel.ListIndex <> 2 Then
    If Len(cboINVPayor) Then
        sqlString = sqlString + " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
    End If
End If

If MultiInvFound = True Then
    MsgBox "Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! ", vbCritical + vbOKOnly
    Exit Sub
End If
    
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim rsSearch As New ADODB.Recordset
Dim strSelFields As String
Dim strSortlist As String
Dim tmpSelCri As String
Dim DataType As String
Dim IsFrst As Boolean
Dim tmpPayor As String
Dim InsDesc As String
Dim Cnt As Integer
IsFrst = True

    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SP_MCOSummaryGSTOP"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, sqlString)
    cmd.Parameters.Append Prm1
    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    rsSearch.CacheSize = 100
    Set rsSearch = cmd.Execute
    Cnt = 1
    Do Until rsSearch.EOF
    
        If IsFrst Then
            If objExcel Is Nothing Then
                Set objExcel = CreateObject("Excel.application")
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MCOInvoice\MCOInvoice.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
            tmpPayor = IIf(Len(rsSearch!PAYCode), rsSearch!PAYCode, "")
            Dim tmpCompName As String
            Dim tmpAdd1 As String
            Dim tmpAdd2 As String
            Dim tmpAdd3 As String
            Dim tmpContact As String
            Dim tmpTelNo As String
            Dim tmpFaxNo As String
            Call PrintPAYInfo(tmpPayor, tmpCompName, tmpAdd1, tmpAdd1, tmpAdd3, tmpContact, tmpTelNo, tmpFaxNo)
            objWorksheet.Cells(8, "A") = tmpCompName
            objWorksheet.Cells(9, "A") = tmpAdd1
            objWorksheet.Cells(10, "A") = tmpAdd2
            objWorksheet.Cells(11, "A") = tmpAdd3
            objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & tmpContact
            objWorksheet.Cells(15, "A") = "TEL :" & " " & tmpTelNo
            objWorksheet.Cells(15, "C") = "FAX :" & " " & tmpFaxNo
            If Trim(rsSearch!MBMCInvoiceNo) <> "" Then
                objWorksheet.Cells(8, "G") = Trim(rsSearch!MBMCInvoiceNo)
            End If
            If cboINVListingVer = "Invoice - by Payor" Then
                tmpPayor = IIf(Len(rsSearch!PAYCode), rsSearch!PAYCode, "")
            Else
                tmpPayor = IIf(Len(rsSearch!GRPCoCode), rsSearch!GRPCoCode, "")
            End If
            IsFrst = False
        End If
        
        If Trim(rsSearch!INSCode) = "I" Then
            InsDesc = "I - INDIVIDUAL"
        ElseIf Trim(rsSearch!INSCode) = "F" Then
            InsDesc = "F - FAMILY"
        ElseIf Trim(rsSearch!INSCode) = "G" Then
            InsDesc = "G - GROUP INDIVIDUAL"
        ElseIf Trim(rsSearch!INSCode) = "H" Then
            InsDesc = "H - GROUP FAMILY"
        End If
        objWorksheet.Cells(20 + CDbl(Cnt), "A") = Cnt
        objWorksheet.Cells(20 + CDbl(Cnt), "B") = Format(rsSearch!MBMBordxDate, "mm/dd/yyyy")
        objWorksheet.Cells(20 + CDbl(Cnt), "C") = InsDesc
        
        If Trim(rsSearch!AdultChild) = "A" Then
            objWorksheet.Cells(20 + CDbl(Cnt), "D") = "Adult"
        ElseIf Trim(rsSearch!AdultChild) = "C" Then
            objWorksheet.Cells(20 + CDbl(Cnt), "D") = "Child"
        End If
        objWorksheet.Cells(20 + CDbl(Cnt), "E") = rsSearch!Empcnt
        objWorksheet.Cells(20 + CDbl(Cnt), "F") = rsSearch!DepCnt
        objWorksheet.Cells(20 + CDbl(Cnt), "G") = rsSearch!BasicMCO
        objWorksheet.Cells(20 + CDbl(Cnt), "H") = IIf(Len(rsSearch!GST), rsSearch!GST, 0)
        Cnt = Cnt + 1
        rsSearch.MoveNext
    Loop
        
    objWorksheet.Range("E43").Formula = "=SUM(E21:E41)"
    objWorksheet.Range("F43").Formula = "=SUM(F21:F41)"
    objWorksheet.Range("G43").Formula = "=SUM(G21:G41)"
    objWorksheet.Range("H43").Formula = "=SUM(H21:H41)"
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Screen.MousePointer = vbDefault
    
    'Code To Details in Excel.
    
    Dim MCOdetails As New ADODB.Recordset
    Dim Parm1 As New ADODB.Parameter
    Dim McoDetCmd As New ADODB.Command
    Dim Rcnt As Integer
    
    IsFrst = True
    Set McoDetCmd.ActiveConnection = MediConnCISDB
    McoDetCmd.CommandText = "SP_MCOGSTInvoiceOP"
    McoDetCmd.CommandType = adCmdStoredProc
    McoDetCmd.CommandTimeout = 300
    Set Prm1 = McoDetCmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, sqlString)
    McoDetCmd.Parameters.Append Prm1
    MCOdetails.CursorLocation = adUseClient
    MCOdetails.CursorType = adOpenForwardOnly
    MCOdetails.CacheSize = 100
    Set MCOdetails = McoDetCmd.Execute
    Rcnt = 4
    Cnt = 1
    Do Until MCOdetails.EOF
        With MCOdetails
        If IsFrst Then
            Set objWorksheet = objWorkbook.Sheets(2)
            Set objWorksheet = objWorkbook.Worksheets.Item(2)
            objExcel.DisplayAlerts = False
            objWorksheet.Cells(Rcnt, "A") = "No"
            objWorksheet.Cells(Rcnt, "B") = "Payor Code"
            objWorksheet.Cells(Rcnt, "C") = "MBM Number"
            objWorksheet.Cells(Rcnt, "D") = "Cover ID"
            objWorksheet.Cells(Rcnt, "E") = "Relationship"
            objWorksheet.Cells(Rcnt, "F") = "Bordx Date"
            objWorksheet.Cells(Rcnt, "G") = "Policy Number"
            objWorksheet.Cells(Rcnt, "H") = "Health Type"
            objWorksheet.Cells(Rcnt, "I") = "Batch No"
            objWorksheet.Cells(Rcnt, "J") = "Name"
            objWorksheet.Cells(Rcnt, "K") = "Group Company"
            objWorksheet.Cells(Rcnt, "L") = "Invoice Number"
            objWorksheet.Cells(Rcnt, "M") = "Basic MCO"
            objWorksheet.Cells(Rcnt, "N") = "GST"
            Rcnt = Rcnt + 1
            IsFrst = False
        End If
        
            Rcnt = Rcnt + 1
            objWorksheet.Cells(Rcnt, "A") = Cnt
            objWorksheet.Cells(Rcnt, "B") = IIf(Len(!PAYCode), !PAYCode, "")
            objWorksheet.Cells(Rcnt, "C") = IIf(Len(!MBMCNUmber), !MBMCNUmber, "")
            objWorksheet.Cells(Rcnt, "D") = IIf(Len(!MBMCCoverID), !MBMCCoverID, "")
            objWorksheet.Cells(Rcnt, "E") = IIf(Len(!MBMCRELCode), !MBMCRELCode, "")
            objWorksheet.Cells(Rcnt, "F") = IIf(Len(!MBMBordxDate), !MBMBordxDate, "")
            objWorksheet.Cells(Rcnt, "G") = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
            objWorksheet.Cells(Rcnt, "H") = IIf(Len(!HLTCode), !HLTCode, "")
            objWorksheet.Cells(Rcnt, "I") = IIf(Len(!MBMBatchNo), !MBMBatchNo, "")
            objWorksheet.Cells(Rcnt, "J") = IIf(Len(!MBMName), !MBMName, "")
            objWorksheet.Cells(Rcnt, "K") = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
            objWorksheet.Cells(Rcnt, "L") = IIf(Len(!MBMCInvoiceNo), !MBMCInvoiceNo, "")
            objWorksheet.Cells(Rcnt, "M") = IIf(Len(!BasicMCO), !BasicMCO, "")
            objWorksheet.Cells(Rcnt, "N") = IIf(Len(!GST), !GST, "")
            Cnt = Cnt + 1
        End With
       MCOdetails.MoveNext
    Loop
    
    
         Dim fPDF As String
         Dim DirPath As String
            If Len(Dir("C:\GSTMCOReports\", vbDirectory)) = 0 Then
               MkDir "C:\GSTMCOReports\"
            End If
         DirPath = "C:\GSTMCOReports\"
         If Len(txtINVInvoiceNo) Then
            fPDF = DirPath + txtInvoiceNo + "_" + Replace(CStr(DateTime.Time), ":", "_") + ")" + ".xls"
         Else
            fPDF = DirPath + tmpPayor + "_" + Replace(CStr(DateTime.Time), ":", "_") + ")" + ".xls"
         End If
         
         objWorksheet.SaveAs fPDF
         objWorkbook.Saved = True
         objWorkbook.Close
         objExcel.Quit
         Set objWorksheet = Nothing
         Set objExcel = Nothing
         
    Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            If Err.Description <> "" Then
                Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
            End If
            Screen.MousePointer = vbDefault
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
    
    'On Error GoTo HANDLE_ERROR
    CompTrans = False
    objExcel.Interactive = False
    Set objExcel = New Excel.Application
    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    RecNotFound = False
    Call ReadRecord(MCORec)
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
            If Trim(MCORec!MBMCRELCode) = "P" Then
                TotalPrin = TotalPrin + 1
            Else
                TotCovered = TotCovered + 1
            End If
            If Len(MCORec!MCOFee) Then
                TOtMco = TOtMco + IIf(IsNumeric(MCORec!MCOFee), CDbl(MCORec!MCOFee), 0)
            Else
                TOtMco = TOtMco
            End If
             
            'PrevKey = Trim(MCORec!INSCode) & "*" & Format(MCORec!MBMBordxDate, "ddmmyyyy") & "*" & Trim(MCORec!MBMAdultChild)
            Cnt = Cnt + 1
        Else
            If Trim(MCORec!MBMCRELCode) = "P" Then
                TotalPrin = TotalPrin + 1
            Else
                TotCovered = TotCovered + 1
            End If
            
            If Len(MCORec!MCOFee) Then
                TOtMco = TOtMco + IIf(IsNumeric(MCORec!MCOFee), CDbl(MCORec!MCOFee), 0)
            Else
                TOtMco = TOtMco
            End If
            
            'TotCovered = TotCovered + MCORec!MBMQty
            'TOtMco = TOtMco + MCORec!MCOFee
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
            'MediConnCISDB.beginTrans
            'medixmasterconn.beginTrans
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
            'medixmasterconn.CommitTrans
            'MediConnCISDB.CommitTrans
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
    Call ReadRecord(MBMCov)
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
           Set ListMaster = lvwInvoice.ListItems.Add(, , IIf(Len(MBMCov!PAYCode), MBMCov!PAYCode, ""))
            ListMaster.SubItems(1) = Trim(MBMCov!MBMCNUmber)
            ListMaster.SubItems(2) = MBMCov!MBMCCoverID
            ListMaster.SubItems(3) = MBMCov!MBMCRELCode
            ListMaster.SubItems(4) = MBMCov!MBMBordxDate
            ListMaster.SubItems(5) = IIf(Len(MBMCov!MBMPolicyNo), Trim(MBMCov!MBMPolicyNo), "")
            ListMaster.SubItems(7) = IIf(Len(MBMCov!MBMBatchNo), MBMCov!MBMBatchNo, "")
            ListMaster.SubItems(8) = Trim(MBMCov!MBMName)
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
            If Len(Trim(cbosubsidiary)) Then
                If Len(MBMCov!Subsidiary) Then
                    ListMaster.SubItems(12) = IIf(Len(Trim(MBMCov!Subsidiary)), Trim(MBMCov!Subsidiary), "")
                End If
                Else
                ListMaster.SubItems(12) = ""
                End If
           
           'tmpTotSupp = tmpTotSupp + 1
           'lblTotSupp = tmpTotSupp
           
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

Private Function SelectionCriterial(tmpMBMTYpe As String) As String

Const METHOD_NAME = "SelectionCriterial"
On Error GoTo debugs

Dim strListingV As String
Dim StrReprint As String
Dim strCriterial As String
Dim tmpInvoice As String
Dim tmpbordxdate As String
Dim tmpBatchNo As String
Dim formattmpBordxDate As String
Dim formatINVBordxDateTo As String
Dim formatINVBordxDateFr As String
    If tmpMBMTYpe = "P" Then
        tmpbordxdate = "MBMBordxDate"
        tmpBatchNo = "MBMBatchNo"
    Else
        tmpbordxdate = "MBMBordxDate"
        tmpBatchNo = "MBMBatchNo"
    End If
    If cboINVCriteriaSel.ListIndex = 2 Then
        strCriterial = " AND MBMCInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
    ElseIf cboINVCriteriaSel.ListIndex = 0 Then
        strCriterial = " AND (" & tmpbordxdate & " = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                   " AND " & tmpBatchNo & " = " & Trim(txtINVBatchNo) & ") AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
    ElseIf cboINVCriteriaSel.ListIndex = 1 Then
        strCriterial = " AND " & tmpbordxdate & " >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
              " AND " & tmpbordxdate & " <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "' AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '') "
        If Len(cboINVGrpCom) Then
            strCriterial = strCriterial & "AND GRPCoCode ='" & Mid(cboINVGrpCom, 1, 4) & "'"
        If Len(Trim(cbosubsidiary)) Then
        strCriterial = strCriterial + " AND Subsidiary = '" & Trim(cbosubsidiary) & "'"
        End If
            
        End If
        'selection Criteria for Group Company
              
              
    End If
    If cboINVListingVer = "Invoice - by Payor" Then
       strCriterial = strCriterial + " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
    End If
    
    'If cboINVPrintOpt.ListIndex = 0 Then
    '    StrReprint = " AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')"
    'ElseIf cboINVPrintOpt.ListIndex = 1 Then
    '    StrReprint = " AND (MBMCInvoiceNo IS NOT NULL OR MBMCInvoiceNo<> '')"
    'End If
    
    SelectionCriterial = strCriterial & strListingV & StrReprint
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Function
Private Sub ReadRecord(rsSearch As ADODB.Recordset)

Const METHOD_NAME = "ReadRecord"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim strSelFields As String
Dim strSortlist As String
Dim tmpSelCri As String
Dim DataType As String
    tmpSelCri = SelectionCriterial(DataType)
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SP_MCOGSTInvoiceOP"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm2 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, tmpSelCri)
    cmd.Parameters.Append Prm2
    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    rsSearch.CacheSize = 100
    Set rsSearch = cmd.Execute
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
   
End Sub
Private Sub ReadRecordSummary(rsSearch As ADODB.Recordset)
Const METHOD_NAME = "ReadRecordSummary"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim strSelFields As String
Dim strSortlist As String
Dim tmpSelCri As String
Dim DataType As String
    tmpSelCri = SelectionCriterial(DataType)
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SP_MCOSummaryGSTOP"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm2 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, tmpSelCri)
    cmd.Parameters.Append Prm2
    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    rsSearch.CacheSize = 100
    Set rsSearch = cmd.Execute
    
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
Dim strsqlPay As String

    If cboINVListingVer = "Invoice - by Payor" Then
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
        'Get GST Account number
        
        tmpPayGrp.MoveNext
    Loop
    tmpPayGrp.Close
    Set tmpPayGrp = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
