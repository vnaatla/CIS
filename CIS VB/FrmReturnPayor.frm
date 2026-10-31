VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReturnPayor 
   Caption         =   "Bordx Returned From Grp. Company"
   ClientHeight    =   8208
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8208
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   25
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1800
         TabIndex        =   45
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   4080
         TabIndex        =   44
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdToPayor 
         Caption         =   "&Update"
         Height          =   375
         Left            =   8400
         TabIndex        =   18
         Top             =   220
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6960
         TabIndex        =   16
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10800
         TabIndex        =   21
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10200
         TabIndex        =   20
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7680
         TabIndex        =   17
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9240
         TabIndex        =   19
         Top             =   220
         Width           =   975
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   120
      TabIndex        =   23
      Top             =   6360
      Width           =   11655
      Begin MSMask.MaskEdBox TxtDate 
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   195
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
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
         Left            =   480
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Returned From Group Company"
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
         Top             =   120
         Width           =   4605
      End
   End
   Begin MSComctlLib.ListView lstBordxDelivery 
      Height          =   4095
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   7218
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Group Company"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Returned Fr Grp Company Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Fr Grp Company"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "A/C Period"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   120
      TabIndex        =   30
      Top             =   400
      Width           =   11655
      Begin VB.ComboBox CboStatus 
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   930
         Width           =   3375
      End
      Begin VB.TextBox TxtYear1 
         Height          =   285
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtYear2 
         Height          =   285
         Left            =   9360
         MaxLength       =   4
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtMonth1 
         Height          =   285
         Left            =   7440
         MaxLength       =   2
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtMonth2 
         Height          =   285
         Left            =   9360
         MaxLength       =   2
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxAmt1 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1260
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxAmt2 
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1260
         Width           =   1695
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   930
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   487
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Top             =   930
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   487
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Caption         =   "Status"
         Height          =   255
         Left            =   6360
         TabIndex        =   42
         Top             =   1005
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "A/C Year"
         Height          =   255
         Left            =   6360
         TabIndex        =   41
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   9000
         TabIndex        =   40
         Top             =   645
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "A/C Month"
         Height          =   255
         Left            =   6360
         TabIndex        =   39
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   9000
         TabIndex        =   38
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx Amt"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   36
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1000
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   34
         Top             =   1000
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Grp. Company"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   650
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   650
         Width           =   255
      End
   End
End
Attribute VB_Name = "FrmReturnPayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim BordxNoArray() As String
Dim searchCriteria As String
Dim ItemChecked As Boolean
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean
Dim Current2 As String
Dim TotalAmt2 As String
Dim UpdateRecord As Boolean
Private m_blnDirection(1 To 9) As Boolean

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub
Private Sub CboPayor_Click()
    If CboPayor.Text = "" Then
        If InStr(CboPayor.Text, "null") Then
            CboPayor.Enabled = False
        End If
        TxtBordx1.SetFocus
    End If
    TxtBordx1.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
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
Call ActionLogUpdate(12, "FrmReturnPayor")
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    Current2 = 0
    TotalAmt2 = 0
    Check = False
    Check2 = False
    Label14.Visible = False
    TxtBordx2.Visible = False
    ReDim BordxNoArray(0)
    Call ComboListing
    ItemChecked = False
End Sub


'**********************Start Command Button *******************************

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstBordxDelivery.ListItems.Clear
        Label14.Visible = False
        TxtBordx2.Visible = False
        TxtBordx1.SetFocus
    Else
        lstBordxDelivery.ListItems.Clear
        Label14.Visible = True
        TxtBordx2.Visible = True
        TxtBordx1.SetFocus
    End If
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelBordxReturnPayor(lstBordxDelivery)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdClear_Click()
    CboPayor.Clear
    CboStatus.Clear
    Call ClearForm(Me)
    Current2 = 0
    TotalAmt2 = 0
    Check = False
    Check2 = False
    TxtBordx1.SetFocus
    lstBordxDelivery.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    ReDim BordxNoArray(0)
    Call ComboListing
End Sub

Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs

    ItemChecked = False
    LblTotalAmt = ""
    lblTotal = ""
    lblCurrent = ""
    ReDim BordxNoArray(0)
    If ChkScan.Value = 1 Then
        If CboPayor = "" Then
            MsgBox "Please select a Group Company", vbOKOnly + vbCritical, DialogBoxTitle
            CboPayor.SetFocus
        ElseIf TxtBordx1 = "" Then
            MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
            TxtBordx1.SetFocus
            CmdClear.Enabled = True
            CmdToPayor.Enabled = True
            cmdPrint.Enabled = True
            CmdExit.Enabled = True
        Else
            Screen.MousePointer = vbHourglass
            listrecord = lstBordxDelivery.ListItems.Count
            If listrecord <> 0 Then
                For listloop1 = 1 To listrecord
                    Call BordxNoCheck
                Next listloop1
            End If
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    CmdClear.Enabled = False
                    CmdToPayor.Enabled = False
                    cmdPrint.Enabled = False
                    CmdExit.Enabled = False
                    Call SearchRecord
                End If
                Check = False
                CmdClear.Enabled = True
                CmdToPayor.Enabled = True
                cmdPrint.Enabled = True
                CmdExit.Enabled = True
            Screen.MousePointer = Default
            TxtBordx1 = ""
            TxtBordx1.SetFocus
        End If
        TxtBordx1.SetFocus
    Else
        Screen.MousePointer = vbHourglass
        Call SearchRecord
        Screen.MousePointer = Default
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset
Dim GrpComName As String
    strAppend = ""
    GrpComName = ""
    GrpComName = Replace(CboPayor, "'", "''")
    If Len(Trim(GrpComName)) Then
        strAppend = strAppend + " And GroupCompany = '" & Trim(GrpComName) & "'"
    End If
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " AND BordxRefNo = '" & Trim(TxtBordx1) & "'"
        End If
    Else
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " AND BordxRefNo >= '" & Trim(TxtBordx1) & "'"
        End If
        If Len(Trim(TxtBordx2)) Then
            strAppend = strAppend + " AND BordxRefNo <= '" & Trim(TxtBordx2) & "'"
        End If
    End If
    If IsDate(txtBordxDate1) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    If IsDate(txtBordxDate2) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    If Len(Trim(TxtBordxAmt1)) Then
        strAppend = strAppend + " AND BordxClaimAmt >= '" & Trim(TxtBordxAmt1) & "'"
    End If
    If Len(Trim(TxtBordxAmt2)) Then
        strAppend = strAppend + " AND BordxClaimAmt <= '" & Trim(TxtBordxAmt2) & "'"
    End If
    If Len(Trim(TxtMonth1)) Then
        strAppend = strAppend + " AND PAYToACPeriodMth >= '" & Trim(TxtMonth1) & "'"
    End If
    If Len(Trim(TxtMonth2)) Then
        strAppend = strAppend + " AND PAYToACPeriodMth <= '" & Trim(TxtMonth2) & "'"
    End If
    If Len(Trim(TxtYear1)) Then
        strAppend = strAppend + " AND PAYToACPeriodYR >= '" & Trim(TxtYear1) & "'"
    End If
    If Len(Trim(TxtYear2)) Then
        strAppend = strAppend + " AND PAYToACPeriodYR <= '" & Trim(TxtYear2) & "'"
    End If
    If Mid(CboStatus, 1, 1) = "1" Then
        strAppend = strAppend + " AND BordxPAYDOCFRStatus = '1'"
    End If
    If Mid(CboStatus, 1, 1) = "0" Then
        strAppend = strAppend + " AND BordxPAYDOCFRStatus = '0'"
    End If
    searchCriteria = strAppend
    If ChkScan.Value = 0 Then
        Call BordxDeliveryListing(strAppend)
    Else
        Call BordxDeliveryListing2(strAppend)
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub CmdStop_Click()
    intExit = 1
    TxtBordx1.SetFocus
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdToPayor_Click()
Const METHOD_NAME = "CmdToPayor_Click"
On Error GoTo debugs
Dim ii As Integer
Dim Cnt As Integer
Dim recordFound As Boolean
Dim recordPassed As Boolean
Dim pos1, pos2 As Integer
    If ItemChecked = True Then
        If TxtDate = "__/__/____" Then
            MsgBox "Please Enter in date", vbCritical
            TxtDate.SetFocus
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                Cnt = 0
                recordFound = False
                If lstBordxDelivery.ListItems.Count <> 0 Then
                    For ii = 0 To UBound(BordxNoArray)
                        If Trim(BordxNoArray(ii)) <> "" Then
                            UpdateReturnPayor BordxNoArray(ii)
                            recordFound = True
                        End If
                    Next ii
                    If UpdateRecord = True And recordFound = True Then
                        MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                        Erase BordxNoArray
                        Arraycnt = 0
                        ReDim BordxNoArray(0)
                        lstBordxDelivery.ListItems.Clear
                    End If
                End If
            End If
        End If
    Else
        MsgBox "Please select at least one record to be updated! ", vbCritical
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'**********************End Command Button *******************************

'**********************Start BordxDelivery Listing  *******************************
Private Sub BordxDeliveryListing(Optional strAppend As String)
Const METHOD_NAME = "BordxDeliveryListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim total As String
Dim current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    total = 0
    current = 0
    lstBordxDelivery.ListItems.Clear
    strAppendCase = "8"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
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
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
    Do
        Do Until rst2.EOF
            With rst2
                Set Record = lstBordxDelivery.ListItems.Add(, , "")
                If Len(!BordxPAYCode) Then
                    Record.SubItems(1) = !BordxPAYCode
                End If
                If Len(!GroupCompany) Then
                    Record.SubItems(2) = Trim(!GroupCompany)
                End If
                If Len(!Bordxrefno) Then
                    Record.SubItems(3) = !Bordxrefno
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(4) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If Len(!BordxClaimAmt) Then
                    Record.SubItems(5) = Format(!BordxClaimAmt, "#,##0.00")
                End If
                If Len(!BordxPayDocFRStatus) Then
                   Record.SubItems(6) = !BordxPayDocFRStatus
                End If
                If Len(!BordxPAYRecDocDate) Then
                    Record.SubItems(7) = Format(!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                If Len(!PAYToACPeriodMth & "/" & !PAYToACPeriodYR) Then
                    Record.SubItems(8) = Trim(!PAYToACPeriodMth & "/" & !PAYToACPeriodYR)
                End If
                current = current + 1
                lblCurrent = current
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt = TotalAmt + rst2!BordxClaimAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
        
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub BordxDeliveryListing2(Optional strAppend As String)
Const METHOD_NAME = "BordxDeliveryListing2"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
    strAppendCase = "8"
    sql = sql + strAppend
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
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
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
    Do
        Do Until rst2.EOF
            With rst2
                Set Record = lstBordxDelivery.ListItems.Add(, , "")
                If Len(!BordxPAYCode) Then
                    Record.SubItems(1) = !BordxPAYCode
                End If
                If Len(!GroupCompany) Then
                    Record.SubItems(2) = Trim(!GroupCompany)
                End If
                If Len(!Bordxrefno) Then
                    Record.SubItems(3) = !Bordxrefno
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(4) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If Len(!BordxClaimAmt) Then
                    Record.SubItems(5) = Format(!BordxClaimAmt, "#,##0.00")
                End If
                If Len(!BordxPayDocFRStatus) Then
                   Record.SubItems(6) = !BordxPayDocFRStatus
                End If
                If Len(!BordxPAYRecDocDate) Then
                    Record.SubItems(7) = Format(!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                If Len(!PAYToACPeriodMth & "/" & !PAYToACPeriodYR) Then
                    Record.SubItems(8) = Trim(!PAYToACPeriodMth & "/" & !PAYToACPeriodYR)
                End If
                Current2 = Current2 + 1
                lblTotal = Current2
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt2 = TotalAmt2 + rst2!BordxClaimAmt
                End If
                LblTotalAmt = Format(TotalAmt2, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'**********************End BordxDelivery Listing *******************************


'**********************Start Check ListView Record*******************************
Private Sub lstBordxDelivery_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "lstBordxDelivery_ItemCheck"
On Error GoTo debugs
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Static Arraycnt As Integer
    If Item.Checked = True Then
        If Item.SubItems(6) = False Then
            Arraycnt = IIf(UBound(BordxNoArray) = 0, 1, UBound(BordxNoArray) + 1)
            ReDim Preserve BordxNoArray(Arraycnt)
            BordxNoArray(Arraycnt) = Item.SubItems(3)
            Arraycnt = Arraycnt + 1
            ItemChecked = True
        Else
            MsgBox "Record has been Updated!", vbCritical
            Item.Checked = False
        End If
    Else
        newbound = 0
        If UBound(BordxNoArray) <> 0 Then
            For i = 0 To UBound(BordxNoArray)
                If (Item.SubItems(3)) <> BordxNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = BordxNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            Erase BordxNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim BordxNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                BordxNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'**********************End Check ListView Record*******************************

'**********************Start Update all Record functions*******************************

Private Function UpdateReturnPayor(tBordxNo As String) As Boolean
Const METHOD_NAME = "UpdateReturnPayor"
On Error GoTo debugs

Dim Update
Dim strsql As String
Dim strsql2 As String
Dim FRPayor As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
Dim ReturnDate As String
    strAppendCase = "9"
    sql = "AND BordxRefNo = '" & Trim(tBordxNo) & "'"
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    FRPayor.CursorLocation = adUseClient
    FRPayor.CursorType = adOpenForwardOnly
    FRPayor.CacheSize = 100
    Set FRPayor = cmd.Execute
    
  If FRPayor.EOF = False Then
    If CDate(FRPayor!BordxPAYSentDate) > CDate(TxtDate) Then
        MsgBox "Date Received from Payor cannot be ealier than Date Sent to Payor!", vbCritical
        TxtDate.SetFocus
        UpdateReturnPayor = False
        UpdateRecord = False
    Else
      ReturnDate = Format(TxtDate, "YYYY/MM/DD")
      strsql2 = " Update Bordx set " & _
               " BordxPayDocFRStatus = 1, BordxPAYRecDocDate = '" & ReturnDate & "'" & _
               " WHERE BordxRefNo = '" & Trim(FRPayor!Bordxrefno) & "'"
      MediConnCISWS.Execute strsql2
      UpdateReturnPayor = True
      UpdateRecord = True
    End If
  End If
  FRPayor.Close
  Set FRPayor = Nothing
  
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'**********************End Update all Record functions*******************************

'**********************Start Sorted ListView *******************************
Private Sub lstBordxDelivery_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'**********************Start Combo Listing*******************************
Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs

Dim PAY As New ADODB.Recordset
Dim monthstart
Dim yearstart
        
    PAY.Open "Select DISTINCT GRPCoName from GRPCompany", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until PAY.EOF
        CboPayor.AddItem Trim(PAY!GrpCoName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    With CboStatus
    .AddItem "0 - Not Returned"
    .AddItem "1 - Returned"
    .AddItem "2 - Both"
    .Text = "0 - Not Returned"
    End With
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
'**********************End Combo Listing*******************************

Private Sub TxtBordxAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBordxAmt2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBordxDate1_LostFocus()
    If IsDate(txtBordxDate1) Then
    Else
        If txtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtBordxDate2_LostFocus()
    If IsDate(txtBordxDate2) Then
    Else
        If txtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDate_LostFocus()
    If IsDate(TxtDate) Then
    Else
        If TxtDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDate.SetFocus
        End If
    End If
End Sub

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim R As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstBordxDelivery.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        R = SendMessageAny(lstBordxDelivery.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Public Sub SetCheck(hwnd As Long, lItemIndex As Long, bState As Boolean)
Dim LV As LVITEM
Dim R As Long
    With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
    End With
    R = SendMessageAny(hwnd, LVM_SETITEMSTATE, lItemIndex, LV)
End Sub

Private Function BordxNoCheck()
    If Trim(ChkStr(TxtBordx1)) = ChkStr(Trim(lstBordxDelivery.ListItems(listloop1).SubItems(2))) Then
        Check = True
    End If
End Function
