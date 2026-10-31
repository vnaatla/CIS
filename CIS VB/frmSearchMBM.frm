VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchMBM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search For Member"
   ClientHeight    =   7404
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9348
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7404
   ScaleWidth      =   9348
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   6360
      Width           =   9255
      Begin VB.TextBox totalSuppCount 
         Height          =   285
         Left            =   5040
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TotalPrincipalCount 
         Height          =   285
         Left            =   3480
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
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
         Left            =   7560
         TabIndex        =   9
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   6120
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   8280
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   315
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         Caption         =   "Supp"
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Principal"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame10 
      Height          =   810
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   60
         TabIndex        =   0
         Top             =   240
         Width           =   3705
      End
      Begin VB.ComboBox CboDataStatus 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label67 
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
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Data Status"
         Height          =   255
         Left            =   5880
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Height          =   810
      Left            =   4800
      TabIndex        =   11
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton OptPolSubNo 
         Caption         =   "Policy Sub Num"
         Height          =   240
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   1545
      End
      Begin VB.OptionButton OptEmployee 
         Caption         =   "Employee Num"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1425
      End
      Begin VB.OptionButton txtMBMName6 
         Caption         =   "Mem Name"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton txtMBMPolicyNo1 
         Caption         =   "Policy Num"
         Height          =   240
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton txtMBMNo1 
         Caption         =   "Mem Num"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.OptionButton txtIcBcPc1 
         Caption         =   "IC Num"
         Height          =   240
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   825
      End
   End
   Begin MSComctlLib.ListView lstMBMInquiryList 
      Height          =   5565
      Left            =   0
      TabIndex        =   20
      Top             =   840
      Width           =   9255
      _ExtentX        =   16320
      _ExtentY        =   9821
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IC No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "PolSt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "DataSt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Eff Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Exp Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Group Company"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Dept"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmSearchMBM"
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

Private Sub cmdClear_Click()
txtSearch.Text = ""
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
    Call ActionLogUpdate(12, "frmSearchMBM")
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intExit = 1
End Sub

Private Sub MBMInquiryList()
Const METHOD_NAME = "MBMInquiryList"
'On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim GrpCompany As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm2 As New ADODB.Parameter
    
      strCriteria = ""
              
      If Len(Trim(txtSearch)) Then
         If txtMBMNo1.Value = True Then
             strCriteria = strCriteria & " AND MBMNumber =  '" & RemStr(txtSearch) & "'"
         ElseIf txtMBMPolicyNo1.Value = True Then
             strCriteria = strCriteria & " AND MBMPolicyNo = '" & RemStr(txtSearch) & "'"
         ElseIf txtMBMName6.Value = True Then
             strCriteria = strCriteria & " AND MBMName  LIKE '%" & RemStr(txtSearch) & "%' "
             'strCriteria = strCriteria & " AND MBMName LIKE '" & RemStr(txtSearch) & "%' "
         ElseIf txtIcBcPc1.Value = True Then
             If Len(txtSearch) >= 12 Then
                Dim LPosition As Integer
                Dim StrIc1 As String
                Dim StrIc2 As String
                LPosition = InStr(txtSearch, "-")
                If LPosition >= 1 Then
                    StrIc1 = txtSearch
                    StrIc2 = Replace(txtSearch, "-", "")
                Else
                    StrIc1 = txtSearch
                    StrIc2 = Mid(txtSearch, 1, 6) & "-" & Mid(txtSearch, 7, 2) & "-" & Mid(txtSearch, 9, 4)
                End If
                strCriteria = strCriteria & " AND ( MBMIcBcPp = '" & RemStr(StrIc1) & "'  or  MBMIcBcPp = '" & RemStr(StrIc2) & "')"
           Else
                strCriteria = strCriteria & " AND MBMIcBcPp = '" & RemStr(txtSearch) & "'"
           End If
         ElseIf OptEmployee.Value = True Then
             strCriteria = strCriteria & " AND MBMEmployeeNo =  '" & RemStr(txtSearch) & "'"
         ElseIf OptPolSubNo.Value = True Then
             strCriteria = strCriteria & " AND MBMPOLSubNo1 =  '" & RemStr(txtSearch) & "'"
         End If
      End If
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Search_MBMCLN2013"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
    cmd.Parameters.Append Prm1
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
            
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
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
    If txtMBMName6.Value = True Then
        strCriteria = ""
        strCriteria = strCriteria & " AND MBMName LIKE '%" & RemStr(txtSearch) & "%'"
    
        Set cmd2.ActiveConnection = medixmasterconn
        cmd2.CommandText = "Search_MBMCLN2013"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm2 = cmd2.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
        cmd2.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
                
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
        intExit = 0
        rst2.Close
        Set rst2 = Nothing
    End If
'debugs:
'If Err.Description <> "" Then
'    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
'End If
End Sub

Private Sub SuppInquiryList()
Const METHOD_NAME = "SuppInquiryList"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim GrpCompany As String
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm2 As New ADODB.Parameter
Dim cmd3 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter




      strCriteria = ""
        
      If Len(Trim(txtSearch)) Then
         If txtMBMNo1.Value = True Then
             strCriteria = strCriteria & " AND MBMCNumber =  '" & RemStr(txtSearch) & "'"
         ElseIf txtMBMName6.Value = True Then
            strCriteria = strCriteria & " AND MBMCName LIKE '%" & RemStr(txtSearch) & "%' "
             'strCriteria = strCriteria & " AND MBMCName LIKE '%" & RemStr(txtSearch) & "%' "
         ElseIf txtIcBcPc1.Value = True Then
             strCriteria = strCriteria & " AND MBMCIcBcPpNo =  '" & RemStr(txtSearch) & "'"
         End If
      End If
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "Search_MBMCovCLN2013"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 300
    Set Prm2 = cmd1.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
    cmd1.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd1.Execute
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
                If Len(!MBMPolExpDate) Then
                   Record.SubItems(7) = Trim(!MBMPolExpDate)
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
    rst2.Close
    Set rst2 = Nothing
    
    If txtMBMName6.Value = True Then
        strCriteria = ""
        strCriteria = strCriteria & " AND MBMCName LIKE '%" & RemStr(txtSearch) & "%' "
        Set cmd3.ActiveConnection = medixmasterconn
        cmd3.CommandText = "Search_MBMCovCLN2013"
        cmd3.CommandType = adCmdStoredProc
        cmd3.CommandTimeout = 300
        Set Prm3 = cmd3.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
        cmd3.Parameters.Append Prm3
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd3.Execute
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
                    If Len(!MBMPolExpDate) Then
                       Record.SubItems(7) = Trim(!MBMPolExpDate)
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
        
        rst2.Close
        Set rst2 = Nothing
    End If
    
    If lstMBMInquiryList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
   ' End If
debugs:
'If Err.Description <> "" Then
'    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
'End If
End Sub

' ######### Button Clicking
Private Sub CmdExit_Click()
    intExit = 1
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo ErrorShow
    
    Screen.MousePointer = vbHourglass
    lstMBMInquiryList.ListItems.Clear
    lblCurrent = ""
    TotalPrincipalCount = ""
    totalSuppCount = ""
    Current = 0
    cmdStop.Enabled = True
    If Len(txtSearch) >= 3 Then
            CmdClear.Enabled = False
            cmdExit.Enabled = False
            Call MBMInquiryList
            If txtMBMName6.Value = True Or txtMBMNo1.Value = True Or txtIcBcPc1.Value = True Then
                Call SuppInquiryList
            End If
            If lstMBMInquiryList.ListItems.Count = 0 Then
                MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
            End If
    Else
        MsgBox "Search String must be at least 3 characters!", vbInformation
        txtSearch.SetFocus
        cmdStop.Enabled = False
    End If
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdStop.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorShow:
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdStop.Enabled = True
    lstMBMInquiryList.Enabled = True
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
    'If Err.Description <> "" Then
    '    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
    'End If
End Sub

Private Sub CmdStop_Click()
    ' stop
    intExit = 1
    CmdClear.Enabled = True
    cmdExit.Enabled = True
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
   
End Sub

Private Sub PlnDesc(PCode, GrpCompany As String)
Const METHOD_NAME = "PlnDesc"
On Error GoTo debugs

Dim PLNRec As New ADODB.Recordset
Dim PlnSql As String
    PlnSql = "SELECT PLNCode, GRPCompany From PLN WHERE PLNCode='" & Trim(PCode) & "'"
    PLNRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not PLNRec.EOF
        If Trim(PLNRec!GrpCompany) <> "" Then
            GrpCompany = PLNRec!GrpCompany
        End If
        PLNRec.MoveNext
    Loop
    PLNRec.Close
    Set PLNRec = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
