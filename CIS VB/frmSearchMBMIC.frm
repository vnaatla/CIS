VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchMBMIC 
   Caption         =   "Search MBM"
   ClientHeight    =   6564
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9612
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleMode       =   0  'User
   ScaleWidth      =   9645
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   9615
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   255
         Left            =   8520
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   255
         Left            =   7440
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame10 
      Height          =   690
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox CboDataStatus 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   2
         Top             =   570
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox cboEXPStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   1
         Top             =   900
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label txtSearch 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Search"
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Data Status"
         Height          =   255
         Left            =   5880
         TabIndex        =   6
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label72 
         Caption         =   "Policy Status"
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
         Left            =   5880
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label11 
         Caption         =   "Exp. Status"
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
         Left            =   5880
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin MSComctlLib.ListView lstMBMInquiryList 
      Height          =   5115
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   9615
      _ExtentX        =   16955
      _ExtentY        =   9017
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member No"
         Object.Width           =   2198
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IC No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Policy No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Eff Date"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Exp Date"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Group Comp"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Dept"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmSearchMBMIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
' To determine what is the source
Public Source As Integer
Public intInvoice As Integer
Public intInvoiceMBM As Integer
Dim intDentInvCln As Integer
Dim intMBMAdj As Integer
Dim intMBMInquiry  As Integer
Dim intMBMTopup As Integer
Dim Current As String
Dim TotalS As String
Dim TotalP As String
Private m_blnDirection(1 To 11) As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub cmdClear_Click()
txtSearch = ""
txtMBMName6.Value = True
lstMBMInquiryList.ListItems.Clear
lblCurrent = ""
txtTotalRecord = ""
TotalPrincipalCount = ""
totalSuppCount = ""
TotalP = 0
TotalS = 0
Current = 0

End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmSearchMBMIC")
    intExit = 0
    intInvoice = 1
    intInvoiceMBM = 2
    intMBMInquiry = 3
    intDentInvCln = 4
    intMBMAdj = 5
    intMBMInquiry = 6
    intMBMTopup = 7
    Current = 0
    total = 0
    Call CenterForm(Me)
    txtSearch = frmInvoice.txtPatientIC
    If Len(Trim(txtSearch)) <> "" Then
        lstMBMInquiryList.Enabled = False
        Call MBMInquiryList
        Call SuppInquiryList
        lstMBMInquiryList.Enabled = True
    Else: Exit Sub
    End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intExit = 1
End Sub

Private Sub MBMInquiryList()
Const METHOD_NAME = "MBMInquiryList"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim GRPCompany As String
      strCriteria = ""
      If Len(Trim(txtSearch)) Then
        strCriteria = strCriteria & " AND MBMIcBcPp =  '" & RemStr(txtSearch) & "'"
      End If
    sqlstr = ""
    sqlstr = " Select MBMNumber, MBMPolicyNo, MBMName, MBMIcBcPp, MBMDepartment, " & _
             " MBMStatus, MBMDataStatus, MBMCCoverID, MBMPayorEffDate, MBMPayorExpDate FROM VIEW_CIS_Search_MBM Where 0 = 0 "
    If Len(Trim(strCriteria)) Then
        sqlstr = sqlstr + strCriteria
    End If
    rst2.Open sqlstr, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    
    Do
        Do While Not rst2.EOF
            TotalP = rst2.RecordCount
            TotalPrincipalCount = TotalP
            With rst2
                Set Record = lstMBMInquiryList.ListItems.Add(, , !MBMNumber)
                If Len(!MBMCCoverID) Then
                    Record.SubItems(1) = !MBMCCoverID
                End If
                If Len(!MBMName) Then
                    Record.SubItems(2) = Trim(!MBMName)
                End If
                Record.SubItems(2) = Trim(!MBMName)
                If Len(!MBMIcBcPp) Then
                   Record.SubItems(3) = Trim(!MBMIcBcPp)
                End If
                If Len(!MBMStatus) Then
                   Record.SubItems(4) = Trim(!MBMStatus)
                End If
                If Len(!MBMDataStatus) Then
                   Record.SubItems(5) = Trim(!MBMDataStatus)
                End If
                If Len(!MBMPayorEffDate) Then
                   Record.SubItems(6) = Trim(!MBMPayorEffDate)
                End If
                If Len(!MBMPayorEXPDate) Then
                   Record.SubItems(7) = Trim(!MBMPayorEXPDate)
                End If
                Current = Current + 1
                lblCurrent = Current
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    'End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SuppInquiryList()
Const METHOD_NAME = "SuppInquiryList"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim GRPCompany As String
    
      strCriteria = ""
        
      If Len(Trim(txtSearch)) Then
             strCriteria = strCriteria & " AND MBMCIcBcPpNo =  '" & RemStr(txtSearch) & "'"
      End If
    
    sqlstr = " Select MBMCNumber, MBMCName, MBMCIcBcPpNo, MBMCLNStatus, MBMCLNDataStatus, " & _
             "  MBMCCoverID, MBMPolEffDate, MBMPolExpDate FROM VIEW_CISMBMSupp_Search Where 0 = 0 "
    If Len(Trim(strCriteria)) Then
        sqlstr = sqlstr + strCriteria
    End If

    rst2.Open sqlstr, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                 
    Do
        Do While Not rst2.EOF
        TotalS = rst2.RecordCount
        totalSuppCount = TotalS
                 
            With rst2
                Set Record = lstMBMInquiryList.ListItems.Add(, , !MBMCNUmber)
                If Len(!MBMCCoverID) Then
                    Record.SubItems(1) = !MBMCCoverID
                End If
                Record.SubItems(2) = Trim(!MBMCName)
                If Len(!MBMCICBCPPNo) Then
                   Record.SubItems(3) = Trim(!MBMCICBCPPNo)
                End If
                If Len(!MBMCLNStatus) Then
                   Record.SubItems(4) = Trim(!MBMCLNStatus)
                End If
                If Len(!MBMCLNDataStatus) Then
                   Record.SubItems(5) = Trim(!MBMCLNDataStatus)
                End If
                If Len(!MBMPolEffDate) Then
                   Record.SubItems(6) = Trim(!MBMPolEffDate)
                End If
                If Len(!MBMPolEXPDate) Then
                   Record.SubItems(7) = Trim(!MBMPolEXPDate)
                End If
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    intExit = 0
    If lstMBMInquiryList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    rst2.Close
    Set rst2 = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub cmdExit_Click()
    intExit = 1
    Unload Me
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstMBMInquiryList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(10)

    End Select
End Sub
'**********************End Sorted ListView *******************************

Private Sub lstMBMInquiryList_DblClick()
Const METHOD_NAME = "lstMBMInquiryList_DblClick"
On Error GoTo debugs

    If lstMBMInquiryList.ListItems.Count Then
        If Source = intInvoice Then
            frmInvoice.txtMBMNo = lstMBMInquiryList.SelectedItem.Text & "-" & lstMBMInquiryList.SelectedItem.SubItems(1)
            Call frmInvoice.CmdGo_Click
        ElseIf Source = intInvoiceMBM Then
            frmInvoiceMBM.txtMBMNo = lstMBMInquiryList.SelectedItem.Text & "-" & lstMBMInquiryList.SelectedItem.SubItems(1)
            Call frmInvoiceMBM.CmdGo_Click
        ElseIf Source = intDentInvCln Then
'            FrmInvDentalCln.txtMBMNo = lstMBMInquiryList.SelectedItem.Text & "-" & lstMBMInquiryList.SelectedItem.SubItems(1)
'            Call FrmInvDentalCln.CmdGo_Click
        ElseIf Source = intMBMAdj Then
            FrmMBMCLNAdj.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
            Call FrmMBMCLNAdj.cmdShow_Click
        ElseIf Source = intMBMInquiry Then
            FrmMBMCLNEnq.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
            Call FrmMBMCLNEnq.cmdShow_Click
        ElseIf Source = intMBMTopup Then
            FrmMBMTopupLmt.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text & "-" & lstMBMInquiryList.SelectedItem.SubItems(1)
            Call FrmMBMTopupLmt.CmdGo_Click
        End If
        Unload Me
   End If
   
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PlnDesc(PCode, GRPCompany As String)
Const METHOD_NAME = "PlnDesc"
On Error GoTo debugs
Dim PlnRec As New ADODB.Recordset
Dim PlnSql As String

    PlnSql = "SELECT PLNCode, GRPCompany From PLN WHERE PLNCode='" & Trim(PCode) & "'"
    PlnRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not PlnRec.EOF
        If Trim(PlnRec!GRPCompany) <> "" Then
            GRPCompany = PlnRec!GRPCompany
        End If
        PlnRec.MoveNext
    Loop
    PlnRec.Close
    Set PlnRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub



