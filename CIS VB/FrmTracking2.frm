VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTracking2 
   Caption         =   "Group Company Clinic Tracking"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8592
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Group Company Clinic Tracking"
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
         TabIndex        =   16
         Top             =   75
         Width           =   3885
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   7320
      Width           =   11775
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7920
         TabIndex        =   9
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   10
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   12
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   11
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   160
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstClinicList 
      Height          =   4560
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Width           =   11745
      _ExtentX        =   20722
      _ExtentY        =   8043
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Group Code"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Group Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clinic Code"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Clinic Name"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   2295
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Width           =   11775
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   6840
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Height          =   2295
         Left            =   8280
         TabIndex        =   23
         Top             =   0
         Width           =   3495
         Begin VB.OptionButton Option4 
            Caption         =   "By Clinic Code (No of Group Company)"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   3135
         End
         Begin VB.OptionButton Option3 
            Caption         =   "By Clinic Code"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "By Group Company Code (No of Clinic)"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   3135
         End
         Begin VB.OptionButton Option1 
            Caption         =   "By Group Company Code"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   2895
         End
      End
      Begin VB.TextBox txtGRPCode 
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
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   0
         Top             =   160
         Width           =   930
      End
      Begin VB.ComboBox CboGrpComName 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   430
         Width           =   5175
      End
      Begin VB.TextBox txtCLNCode 
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
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   2
         Top             =   720
         Width           =   915
      End
      Begin VB.ComboBox CboCLNName 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label Label5 
         Caption         =   "Clinic Name"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Clinic Code"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   820
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Group Com. Name"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Group Com. Code"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmTracking2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AA As Integer
Private intExit As Integer
Private GrpCompany As String
Private Clinics As String
Private m_blnDirection(1 To 4) As Boolean

Private Sub Check1_Click()
    Call GRPCOMCLNComboListing
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = ""
    lstClinicList.ListItems.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
    If lstClinicList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstClinicList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
Dim strAppend As String
    strAppend = ""
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    cmdExit.Enabled = False
    
    Screen.MousePointer = vbHourglass
        Call SearchRecords
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "FrmTracking2")
    'Call GRPCOMCLNComboListing
End Sub

Private Sub lstClinicList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)

    If Option1.Value = True Or Option3.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(4)
        End Select
    
    ElseIf Option2.Value = True Or Option4.Value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstClinicList, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstClinicList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        End Select
    End If

End Sub


Private Sub Option1_Click()
    Call ListViewHeader
End Sub

Private Sub Option2_Click()
    Call ListViewHeader
End Sub

Private Sub Option3_Click()
    Call ListViewHeader
End Sub

Private Sub Option4_Click()
    Call ListViewHeader
End Sub

Private Sub GRPCOMCLNComboListing()
Const METHOD_NAME = "GRPCOMCLNComboListing"
On Error GoTo debugs

    Dim ClinicName As New ADODB.Recordset
    Dim GRPCOMName As New ADODB.Recordset
    CboGrpComName.Clear
    CboCLNName.Clear
    GRPCOMName.Open "GRPCompany", MediConnCISDB, adOpenForwardOnly, adLockBatchOptimistic
    ClinicName.Open "CLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until GRPCOMName.EOF
        CboGrpComName.AddItem Trim(GRPCOMName!GrpCoName)
        GRPCOMName.MoveNext
    Loop
    Do Until ClinicName.EOF
        CboCLNName.AddItem Trim(ClinicName!CLNName)
        ClinicName.MoveNext
    Loop
    ClinicName.Close
    Set ClinicName = Nothing
    GRPCOMName.Close
    Set GRPCOMName = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ListViewHeader()
    
If Option1.Value = True Then
    lstClinicList.ColumnHeaders(1).Text = "Group Code"
    lstClinicList.ColumnHeaders(2).Text = "Group Name"
    lstClinicList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstClinicList.ColumnHeaders(3).Text = "Clinic Code"
    lstClinicList.ColumnHeaders(3).Alignment = lvwColumnLeft
    lstClinicList.ColumnHeaders(4).Text = "Clinic Name"
    lstClinicList.ColumnHeaders(4).Alignment = lvwColumnLeft
ElseIf Option2.Value = True Then
    lstClinicList.ColumnHeaders(1).Text = "Group Code"
    lstClinicList.ColumnHeaders(2).Text = "Group Name"
    lstClinicList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstClinicList.ColumnHeaders(3).Text = "No of Clinic"
    lstClinicList.ColumnHeaders(3).Alignment = lvwColumnRight
    For AA = 4 To 4
        lstClinicList.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf Option3.Value = True Then
    lstClinicList.ColumnHeaders(1).Text = "Clinic Code"
    lstClinicList.ColumnHeaders(2).Text = "Clinic Name"
    lstClinicList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstClinicList.ColumnHeaders(3).Text = "Group Code"
    lstClinicList.ColumnHeaders(3).Alignment = lvwColumnLeft
    lstClinicList.ColumnHeaders(4).Text = "Group Name"
    lstClinicList.ColumnHeaders(4).Alignment = lvwColumnLeft
ElseIf Option4.Value = True Then
    lstClinicList.ColumnHeaders(1).Text = "Clinic Code"
    lstClinicList.ColumnHeaders(2).Text = "Clinic Name"
    lstClinicList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstClinicList.ColumnHeaders(3).Text = "No of Group Company"
    lstClinicList.ColumnHeaders(3).Alignment = lvwColumnRight
    For AA = 4 To 4
        lstClinicList.ColumnHeaders(AA).Text = ""
    Next AA
End If

End Sub

Public Sub ExportListViewtoExcel(lvwList As Control)
Const METHOD_NAME = "ExportListViewtoExcel"
On Error GoTo debugs

    Dim vntHeader As Variant
    Dim vntData As Variant
    Dim X As Long
    Dim Y As Long
    Dim intCol As Integer
    Dim lngRow As Long
    'Get Counts
    intCol = CInt(lvwList.ColumnHeaders.Count - 1)
    lngRow = CLng(lvwList.ListItems.Count - 1)
    ReDim vntHeader(0)
    ReDim vntData(intCol, lngRow)
    'Create Header Array
    For X = 0 To intCol
        ReDim Preserve vntHeader(X)
        vntHeader(X) = lvwList.ColumnHeaders(X + 1).Text
    Next
    'Create Data Array
    For X = 0 To lngRow
        vntData(0, X) = lvwList.ListItems.Item(X + 1).Text
        For Y = 1 To intCol
            vntData(Y, X) = lvwList.ListItems.Item(X + 1).SubItems(Y)
        Next
    Next
    'Create Excel Object
    OpenExcel vntData, vntHeader
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub ExportRecords(vntData As Variant, vntHeader As Variant, ws As Worksheet)
Const METHOD_NAME = "ExportRecords"
On Error GoTo debugs

    Dim lngRow As Long
    Dim intCol As Integer
    Dim varData As Variant
    Dim intStart As Integer
    Dim X As Long
    Dim Y As Long
    'Select all Cells and and set the number
        ws.Cells.Select
        ws.Cells.NumberFormat = "@"
        ws.Cells(1, 1).Select
        lngRow = UBound(vntData, 2) + 2
        intCol = UBound(vntData, 1) + 1
        intStart = 2 'Start from line 2
        'Freeze Row 2
        ws.Rows(2).Select
        ws.Activate
        'ActiveWindow..FreezePanes = True
        'Add Headers
        For X = 1 To intCol
            varData = vntHeader(X - 1)
            ws.Cells(1, X) = CStr(varData)
            ws.Cells(1, X).Font.Bold = True
        Next
        'Add Data
        For Y = 1 To intCol
            For X = intStart To lngRow
                varData = vntData(Y - 1, X - 2)
                If IsNull(varData) Then 'Make sure no null values, Excel will choke
                    'Add 1 to Move down a column
                    ws.Cells(X + 1, Y) = ""
                Else
                    ws.Cells(X + 1, Y) = CStr(varData) 'Convert to String to preserve formatting
                End If
            Next
        Next
        'Resize Columns to Fit
        ws.Columns.AutoFit

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

        
End Sub


Private Sub OpenExcel(vntData As Variant, vntHeader As Variant)
    Const METHOD_NAME = "OpenExcel"
    On Error GoTo Err_OpenExcel
    Dim objExcel As Excel.Application
    Dim objWrkSht As Worksheet
    Dim X As Integer
    'Create Excel Object
    Set objExcel = CreateObject("Excel.Application")
    'Add the Workbook
    objExcel.Workbooks.Add
    '========== Set Template ===============
        'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\DRCRListing\DRCRListing.XLT")
    '========== Set Template ===============
    Set objWrkSht = objExcel.ActiveWorkbook.Sheets(1)
    objExcel.Visible = True
    'Fill the Workbook with data
    ExportRecords vntData, vntHeader, objWrkSht
    objExcel.Interactive = True
    ' Clean up:
    Set objWrkSht = Nothing
    Set objExcel = Nothing
Err_OpenExcel:
    Select Case Err
        Case 0
        Case 439
        MsgBox "You must have Microsoft Excel installed on your PC.", vbCritical, "Application Not Found"
        Case Else
        MsgBox Err & ": " & error, vbCritical, "OpenExcel Error"
    End Select
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function SearchRecords()
Const METHOD_NAME = "SearchRecords"
On Error GoTo debugs
    Dim sql As String
    Dim strAppend As String
    strAppend = ""
    GrpCompany = ""
    Clinics = ""
    If Len(Trim(txtGRPCode)) Then
       strAppend = strAppend + " And  GRPCOMCode like '%" & ChkStr(txtGRPCode) & "%' "
    End If
    If Len(Trim(txtCLNCode)) Then
       strAppend = strAppend + " And  GRPCOMCLNCode like '%" & ChkStr(txtCLNCode) & "%' "
    End If
    If Len(Trim(CboGrpComName)) Then
         GrpCompany = Replace(CboGrpComName, "'", "''")
         strAppend = strAppend + " AND GrpCoName = '" & Trim(GrpCompany) & "'"
    End If
    If Len(Trim(CboCLNName)) Then
         Clinics = Replace(CboCLNName, "'", "''")
         strAppend = strAppend + " AND CLNName = '" & Trim(Clinics) & "'"
    End If
    If Option1.Value = True Then
        Call GRPCOMListing(strAppend)
    ElseIf Option2.Value = True Then
        Call GRPCOMListing2(strAppend)
    ElseIf Option3.Value = True Then
        Call CLNListing(strAppend)
    ElseIf Option4.Value = True Then
        Call CLNListing2(strAppend)
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub GRPCOMListing(Optional strAppend As String)
Const METHOD_NAME = "GRPCOMListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    
    lblCurrent = 0
    lstClinicList.ListItems.Clear
    strAppendCase = "1"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "CLNTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    'sql = " SELECT GRPComCode, GrpCoName, " & _
          " GrpComClnCode, CLNName " & _
          " From VIEW_GrpCom_Clinic Where 0 = 0 "
    'If Len(Trim(strAppend)) Then
        'sql = sql + strAppend
    'Else
        'sql = sql
    'End If
    'rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            With rst2
                If Len(rst2!GRPComCode) Then
                    Set Record = lstClinicList.ListItems.Add(, , Trim(rst2!GRPComCode))
                Else
                    Set Record = lstClinicList.ListItems.Add(, , "Empty")
                End If
                If Len(rst2!GrpCoName) Then
                   Record.SubItems(1) = Trim(rst2!GrpCoName)
                End If
                If Len(rst2!GrpComClnCode) Then
                   Record.SubItems(2) = Trim(rst2!GrpComClnCode)
                End If
                If Len(rst2!CLNName) Then
                   Record.SubItems(3) = Trim(rst2!CLNName)
                End If
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If lstClinicList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub GRPCOMListing2(Optional strAppend As String)
Const METHOD_NAME = "GRPCOMListing2"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    lblCurrent = 0
    lstClinicList.ListItems.Clear
    strAppendCase = "2"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "CLNTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    'sql = " SELECT GRPComCode, GrpCoName, " & _
          " Count(GrpComClnCode) AS TotalCLN " & _
          " From VIEW_GrpCom_Clinic Where 0 = 0 "
    
    'sql2 = " GROUP BY GRPComCode, GrpCoName "
    
    'If Len(Trim(strAppend)) Then
        'sql = sql + strAppend + sql2
    'Else
        'sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            With rst2
                If Len(rst2!GRPComCode) Then
                    Set Record = lstClinicList.ListItems.Add(, , Trim(rst2!GRPComCode))
                Else
                    Set Record = lstClinicList.ListItems.Add(, , "Empty")
                End If
                If Len(rst2!GrpCoName) Then
                   Record.SubItems(1) = Trim(rst2!GrpCoName)
                End If
                If Len(rst2!totalCLN) Then
                   Record.SubItems(2) = Trim(rst2!totalCLN)
                End If
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If lstClinicList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CLNListing(Optional strAppend As String)
Const METHOD_NAME = "CLNListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    lblCurrent = 0
    lstClinicList.ListItems.Clear
    strAppendCase = "1"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "CLNTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    'sql = " SELECT GRPComCode, GrpCoName, " & _
          " GrpComClnCode, CLNName " & _
          " From VIEW_GrpCom_Clinic Where 0 = 0 "
    'If Len(Trim(strAppend)) Then
        'sql = sql + strAppend
    'Else
        'sql = sql
    'End If
    'rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            With rst2
            
                If Len(rst2!GrpComClnCode) Then
                    Set Record = lstClinicList.ListItems.Add(, , Trim(rst2!GrpComClnCode))
                Else
                    Set Record = lstClinicList.ListItems.Add(, , "Empty")
                End If
                If Len(rst2!CLNName) Then
                   Record.SubItems(1) = Trim(rst2!CLNName)
                End If
                If Len(rst2!GRPComCode) Then
                   Record.SubItems(2) = Trim(rst2!GRPComCode)
                End If
                If Len(rst2!GrpCoName) Then
                   Record.SubItems(3) = Trim(rst2!GrpCoName)
                End If
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If lstClinicList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CLNListing2(Optional strAppend As String)
Const METHOD_NAME = "CLNListing2"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    lblCurrent = 0
    lstClinicList.ListItems.Clear
    strAppendCase = "3"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "CLNTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    'sql = " SELECT GrpComClnCode, CLNName, " & _
          " Count(GRPComCode) AS TotalCom " & _
          " From VIEW_GrpCom_Clinic Where 0 = 0 "
    
    'sql2 = " GROUP BY GrpComClnCode, CLNName "
    
    'If Len(Trim(strAppend)) Then
        'sql = sql + strAppend + sql2
    'Else
        'sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            With rst2
                If Len(rst2!GrpComClnCode) Then
                    Set Record = lstClinicList.ListItems.Add(, , Trim(rst2!GrpComClnCode))
                Else
                    Set Record = lstClinicList.ListItems.Add(, , "Empty")
                End If
                If Len(rst2!CLNName) Then
                   Record.SubItems(1) = Trim(rst2!CLNName)
                End If
                If Len(rst2!totalCom) Then
                   Record.SubItems(2) = Trim(rst2!totalCom)
                End If
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If lstClinicList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboCLNName_Change()
    If CboCLNName.Text = "" Then
        If InStr(CboCLNName.Text, "null") Then
           CboCLNName.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNName_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboCLNName_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCLNName.Text, CboCLNName.SelStart) + Chr(KeyAscii) + Mid(CboCLNName.Text, CboCLNName.SelStart + CboCLNName.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboCLNName.ListCount - 1
        If NewText = LCase(Left(CboCLNName.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCLNName.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboCLNName.Text = ValidValue
            CboCLNName.SelStart = 0
            CboCLNName.SelLength = Len(ValidValue)
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboGrpComName_Change()
    If CboGrpComName.Text = "" Then
        If InStr(CboGrpComName.Text, "null") Then
           CboGrpComName.Enabled = False
        End If
    End If
End Sub

Private Sub CboGrpComName_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboGrpComName_KeyPress"
On Error GoTo debugs

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpComName.Text, CboGrpComName.SelStart) + Chr(KeyAscii) + Mid(CboGrpComName.Text, CboGrpComName.SelStart + CboGrpComName.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboGrpComName.ListCount - 1
        If NewText = LCase(Left(CboGrpComName.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpComName.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboGrpComName.Text = ValidValue
            CboGrpComName.SelStart = 0
            CboGrpComName.SelLength = Len(ValidValue)
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
