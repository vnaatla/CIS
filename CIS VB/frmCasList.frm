VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCasList 
   Caption         =   "List of Cases"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8328
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   8328
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   0
      TabIndex        =   18
      Top             =   4680
      Width           =   8295
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   855
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
         Left            =   6120
         TabIndex        =   9
         Top             =   200
         Width           =   960
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
         Left            =   4200
         TabIndex        =   7
         Top             =   200
         Width           =   960
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   315
         Left            =   5160
         TabIndex        =   8
         Top             =   200
         Width           =   960
      End
      Begin VB.CommandButton cmdExit2 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
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
         TabIndex        =   10
         Top             =   200
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "of"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   195
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8295
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
         Left            =   720
         MaxLength       =   60
         TabIndex        =   0
         Top             =   240
         Width           =   3615
      End
      Begin VB.Frame Frame3 
         Height          =   660
         Left            =   4440
         TabIndex        =   16
         Top             =   120
         Width           =   3735
         Begin VB.OptionButton txtMBMNo1 
            Caption         =   "Mem No"
            Height          =   240
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   900
         End
         Begin VB.OptionButton optCASNo1 
            Caption         =   "Case No"
            Height          =   240
            Left            =   2640
            TabIndex        =   5
            Top             =   120
            Width           =   1065
         End
         Begin VB.OptionButton txtMBMPolicyNo1 
            Caption         =   "Policy No"
            Height          =   240
            Left            =   1440
            TabIndex        =   3
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optMBMName1 
            Caption         =   "MBM Name"
            Height          =   240
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Value           =   -1  'True
            Width           =   1365
         End
         Begin VB.OptionButton txtIcBcPc1 
            Caption         =   "IC No"
            Height          =   240
            Left            =   1440
            TabIndex        =   4
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.OptionButton optReject 
         Caption         =   "Repudiated"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optOpen 
         Caption         =   "Open"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   1560
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   8295
      Begin MSComctlLib.ListView lstCaseList 
         Height          =   3675
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   8085
         _ExtentX        =   14266
         _ExtentY        =   6477
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Case Num"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership Num"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CoverID"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Patient Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pol Num"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Visit Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Clinic"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ICNo"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   " "
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCasList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intExit  As Integer
Public Source As Integer
Dim intInvoice As Integer
Dim intINVDent As Integer


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
    optMBMName1.Value = True
    lstCaseList.ListItems.Clear
    lblCurrent = ""
    lblTotal = ""

End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmCasList")
    intInvoice = 1
    intINVDent = 2
    intExit = 0
End Sub

Private Sub cmdExit2_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Const METHOD_NAME = "cmdSearch_Click"
    On Error GoTo debugs
    Screen.MousePointer = vbHourglass
    
    If Len(txtSearch) >= 3 Then
        CmdClear.Enabled = False
        Call ListCases
    Else
        MsgBox "Search String must be at least 3 characters!", vbInformation
        txtSearch.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    cmdExit2.Enabled = True
End Sub

' ######## List out all the cases which have been discharged
Private Sub ListCases()
Const METHOD_NAME = "ListCases"
On Error GoTo debugs
'On Error Resume Next
Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset
Dim CLNRec As New ADODB.Recordset
Dim CLNSql As String

cmdExit2.Enabled = False

    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
    strsql = "SELECT CASNumber, MBMNumber, MBMPolicyNo, MBMName, MBMCName," & _
             " MBMIcBcPp,  ClinicCode, DateVisit, CoverID, CASStatus " & _
             "  from VIEW_Cas_MBM Where MBMCLNDataStatus <> 'D' " & _
             " AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')"

    If Source = intInvoice Then
        strsql = strsql & " AND CaseType='C'"
    ElseIf Source = intINVDent Then
        strsql = strsql & " AND CaseType='D'"
    End If
             
    If Len(Trim(txtSearch)) Then
        If txtMBMNo1.Value = True Then
            strsql = strsql & " AND ( MBMNumber LIKE  '" & RemStr(txtSearch) & "%' "
            strsql = strsql & " OR MBMCNumber LIKE  '" & RemStr(txtSearch) & "%') "
        ElseIf txtMBMPolicyNo1.Value = True Then
            strsql = strsql & " AND MBMPolicyNo like '" & RemStr(txtSearch) & "%' "
        ElseIf optMBMName1.Value = True Then
            strsql = strsql & " AND MBMName LIKE '%" & RemStr(txtSearch) & "%' "
        ElseIf txtIcBcPc1.Value = True Then
            strsql = strsql & " AND  MBMIcBcPp like  '" & RemStr(txtSearch) & "%' "
        ElseIf optCASNo1.Value = True Then
            strsql = strsql & " AND  CASNumber LIKE  '" & RemStr(txtSearch) & "%' "
        End If
    End If
            
    CAS.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If CAS.RecordCount = 0 Then
        lstCaseList.ListItems.Clear
        MsgBox "No Record Available", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Screen.MousePointer = vbHourglass
        lstCaseList.ListItems.Clear
        lblCurrent = 0
        lblTotal = CAS.RecordCount
        Do
            Do Until CAS.EOF
                
                Set CASList = lstCaseList.ListItems.Add(, , CAS!CASNumber)
                    CASList.SubItems(1) = CAS!MBMNumber
                    CASList.SubItems(2) = CAS!coverid
                    If CAS!coverid <> "00" Then
                        If CAS!MBMCName <> "" Then
                            CASList.SubItems(3) = CAS!MBMCName
                        Else
                            CASList.SubItems(3) = "Member may has been Deleted!"
                        End If
                        
                    Else
                        If Trim(CAS!MBMName) <> "" Then
                            CASList.SubItems(3) = CAS!MBMName
                        Else
                            CASList.SubItems(3) = "Member may has been Deleted!"
                        End If
                    End If
                    
                    CASList.SubItems(4) = IIf(Len(Trim(CAS!MBMPolicyNo)), Trim(CAS!MBMPolicyNo), "")
                    If Len(Trim(CAS!DateVisit)) Then
                        CASList.SubItems(5) = CAS!DateVisit
                    End If
                    
                    If Trim(CAS!CASStatus) <> "" Then
                        CASList.SubItems(6) = CAS!CASStatus
                    End If
                    
                    If Len(Trim(CAS!ClinicCode)) Then
                        CASList.SubItems(7) = CAS!ClinicCode
                        CLNSql = "SELECT CLNCode, CLNName FROM CLN WHERE (CLNPanel = 'C' OR CLNPanel = 'Y')" & _
                                " AND (CLNCode ='" & Trim(CAS!ClinicCode) & "')"

                        CLNRec.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If CLNRec.RecordCount Then
                            CASList.SubItems(7) = Trim(CAS!ClinicCode) & " - " & Trim(CLNRec!CLNName)
                        End If
                        CLNRec.Close
                        Set CLNRec = Nothing
                    End If
                    If Trim(CAS!MBMIcBcPp) <> "" Then
                        CASList.SubItems(8) = CAS!MBMIcBcPp
                    End If
                    
                    lblCurrent = lblCurrent + 1
                    CAS.MoveNext
                    
                DoEvents
                If intExit = 1 Then
                    Check = False
                    Exit Do
                End If
            Loop
        Loop Until Check = False
        Screen.MousePointer = vbDefault
    End If
    intExit = 0
    CAS.Close
    Set CAS = Nothing
    CmdClear.Enabled = True
    cmdExit2.Enabled = True
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstCaseList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Const METHOD_NAME = "lstCaseList_ColumnClick"
    On Error GoTo debugs
    lstCaseList.SortKey = ColumnHeader.Index - 1
    lstCaseList.Sorted = True
    If lstCaseList.SortOrder = lvwAscending Then
        lstCaseList.SortOrder = lvwDescending
    Else
        lstCaseList.SortOrder = lvwAscending
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstCaseList_DblClick()
Const METHOD_NAME = "lstCaseList_DblClick"
On Error GoTo debugs
Dim ctl As Form
    
    If lstCaseList.ListItems.Count <> 0 Then
        
        If Source = intInvoice Then
            Set ctl = frmInvoiceAdj
            With ctl
                .txtCaseNo = lstCaseList.SelectedItem.Text
                 Call .CmdGo_Click
            End With
        
        End If
        'If Source = intINVDent Then
        '    Set ctl = FrmInvDentalAdj
        '    With ctl
        '        .txtCaseNo = lstCaseList.SelectedItem.Text
        '         Call .CmdGo_Click
        '    End With
       '
       ' End If
        Unload Me
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
