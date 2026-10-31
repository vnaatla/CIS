VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxWks 
   Caption         =   "Bordx & Worksheet"
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   37
      Top             =   7080
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   4320
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7800
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6960
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9600
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   8640
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1320
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3480
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Bordx and Worksheet"
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
         TabIndex        =   36
         Top             =   120
         Width           =   3285
      End
   End
   Begin MSComctlLib.ListView lstBordxTrackList 
      Height          =   5175
      Left            =   120
      TabIndex        =   40
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   9123
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Payor"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Group Company"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bordx Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Bordx Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Worksheet No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Worksheet Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Worksheet Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "DFGrp Company (Money)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Receipt No (Grp Company)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11655
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   7800
         MaxLength       =   20
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   9840
         MaxLength       =   20
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox ChkDFMNRB 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox ChkDFP 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin MSMask.MaskEdBox TxtDFMNRB2 
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFMNRB1 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtClaimNo1 
         Height          =   285
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   11
         Top             =   500
         Width           =   1575
      End
      Begin VB.TextBox TxtRecPAY1 
         Height          =   285
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtRecMNRB1 
         Height          =   285
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtClaimNo2 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   12
         Top             =   500
         Width           =   1575
      End
      Begin VB.TextBox TxtRecPAY2 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtRecMNRB2 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox CboGrpCompany 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   4695
      End
      Begin MSMask.MaskEdBox TxtDFP1 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   765
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFP2 
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   765
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   550
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No from"
         Height          =   255
         Left            =   6360
         TabIndex        =   33
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   9480
         TabIndex        =   32
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Left            =   9480
         TabIndex        =   31
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Claim No from"
         Height          =   255
         Left            =   6360
         TabIndex        =   30
         Top             =   550
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Date Fr MNRB"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Date Fr Grp Com"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   26
         Top             =   885
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Rec. No (MNRB)fr."
         Height          =   255
         Left            =   6360
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   9480
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   255
         Left            =   9480
         TabIndex        =   23
         Top             =   795
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Rec. No (GrpCom)"
         Height          =   255
         Left            =   6360
         TabIndex        =   22
         Top             =   795
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmBordxWks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public searchCriteria As String
Private m_blnDirection(1 To 10) As Boolean

Private Sub CboGrpCompany_Change()
Const METHOD_NAME = "CboGrpCompany_Change"
On Error GoTo debugs
If CboGrpCompany.Text = "" Then
    If InStr(CboGrpCompany.Text, "null") Then
       CboGrpCompany.Enabled = False
    End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboGrpCompany_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboGrpCompany_KeyPress"
On Error GoTo debugs
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpCompany.Text, CboGrpCompany.SelStart) + Chr(KeyAscii) + Mid(CboGrpCompany.Text, CboGrpCompany.SelStart + CboGrpCompany.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboGrpCompany.ListCount - 1
 
        If NewText = LCase(Left(CboGrpCompany.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpCompany.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboGrpCompany.Text = ValidValue
                CboGrpCompany.SelStart = 0
                CboGrpCompany.SelLength = Len(ValidValue)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Const METHOD_NAME = "Form_KeyDown"
    On Error GoTo debugs
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdPrintFile_Click()
Const METHOD_NAME = "CmdPrintFile_Click"
On Error GoTo debugs
    Screen.MousePointer = vbHourglass
        Call ExportToExcelBordx(lstBordxTrackList)
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmBordxWks")
    intExit = 0
    lblCurrent = 0
    'lblTotal = 0
    LblTotalAmt = 0
    'Call ComboListing
    Call PayorComboListing
End Sub

'**********************Start Command Button *******************************

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    'cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstBordxTrackList.ListItems.Clear
    lblCurrent = 0
    'lblTotal = 0
    LblTotalAmt = 0
    Call PayorComboListing
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    
    CmdClear.Enabled = False
    CmdPrintFile.Enabled = False
    CmdExit.Enabled = False
    Screen.MousePointer = vbHourglass
        Call SearchRecord
    Screen.MousePointer = Default
    
End Sub
'**********************End Command Button *******************************

'**********************Start Search Record *******************************

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs
Dim strsql As String
Dim StrAppend As String
Dim Sql As String
Dim CAS As New ADODB.Recordset
Dim GrpCompany As String

    StrAppend = ""
    GrpCompany = ""
    
    GrpCompany = Replace(CboGrpCompany, "'", "''")
    
    If Len(Trim(GrpCompany)) Then
        StrAppend = StrAppend + " And GroupCompany = '" & Trim(GrpCompany) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
        StrAppend = StrAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtBordx1)) Then
        StrAppend = StrAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
        StrAppend = StrAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
    End If
    
    If Len(Trim(TxtClaimNo1)) Then
        StrAppend = StrAppend + " AND CASNumber >= '" & RemStr(Trim(TxtClaimNo1)) & "'"
    End If
    
    If Len(Trim(TxtClaimNo2)) Then
        StrAppend = StrAppend + " AND CASNumber <= '" & RemStr(Trim(TxtClaimNo2)) & "'"
    End If
    
    If Len(Trim(TxtRecPAY1)) Then
        StrAppend = StrAppend + " AND ReceiptNoPAY >= '" & RemStr(Trim(TxtRecPAY1)) & "'"
    End If
    
    If Len(Trim(TxtRecPAY2)) Then
        StrAppend = StrAppend + " AND ReceiptNoPAY <= '" & RemStr(Trim(TxtRecPAY2)) & "'"
    End If
    
    'If Len(Trim(TxtRecMNRB1)) Then
        'StrAppend = StrAppend + " AND ReceiptNoMNRB >= '" & RemStr(Trim(TxtRecMNRB1)) & "'"
    'End If
    
    'If Len(Trim(TxtRecMNRB2)) Then
        'StrAppend = StrAppend + " AND ReceiptNoMNRB <= '" & RemStr(Trim(TxtRecMNRB2)) & "'"
    'End If
          
    If IsDate(TxtDFP1) = True Then
        StrAppend = StrAppend + " And BordxPAYRecDate >= '" & Format(TxtDFP1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(TxtDFP2) = True Then
        StrAppend = StrAppend + " And BordxPAYRecDate <= '" & Format(TxtDFP2, "mm/dd/yyyy") & "'"
    End If
    
    'If IsDate(TxtDFMNRB1) = True Then
        'StrAppend = StrAppend + " And BordxMNRBRecDate >= '" & Format(TxtDFMNRB1, "mm/dd/yyyy") & "'"
    'End If
      
    'If IsDate(TxtDFMNRB2) = True Then
        'StrAppend = StrAppend + " And BordxMNRBRecDate <= '" & Format(TxtDFMNRB2, "mm/dd/yyyy") & "'"
    'End If
    
    If ChkDFP.Value = 1 Then
        Sql = ""
        Sql = " AND (BordxPAYRecDate Is null OR BordxPAYRecDate = '') " & ""
        StrAppend = StrAppend + Sql
    End If
    
    'If ChkDFMNRB.Value = 1 Then
        'Sql = ""
        'Sql = " AND (BordxMNRBRecDate Is null OR BordxMNRBRecDate = '') " & ""
        'StrAppend = StrAppend + Sql
    'End If
    
    
    searchCriteria = StrAppend
    Call BordxTrackingListing(StrAppend)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'********************** End Search Record *******************************

'**********************Start BordxTracking Listing *******************************

Private Sub BordxTrackingListing(Optional StrAppend As String)
Const METHOD_NAME = "BordxTrackingListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String


    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    Total = 0
    Current = 0
    
    lstBordxTrackList.ListItems.Clear
            
    strAppendCase = "15"
    
    Sql = Sql + StrAppend
            
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, Sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    'Sql = " Select CASNumber, BordxPAYCode, BordxClaimAmt, " & _
          " BordxDate, BordxRefNo, CLMTotalApproved, ClaimOpendate, " & _
          " BordxPAYRecDate, ReceiptNoPAY, GroupCompany " & _
          " From VIEW_CISBordx_Wks where 0 = 0 "
    
    'If Len(Trim(StrAppend)) Then
        'Sql = Sql + StrAppend
    'End If

    'rst2.Open Sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
                        
            With rst2
                If Len(!BordxPAYCode) Then
                    Set Record = lstBordxTrackList.ListItems.Add(, , !BordxPAYCode)
                End If
                If Len(!GroupCompany) Then
                    Record.SubItems(1) = Trim(!GroupCompany)
                End If
                If Len(!Bordxrefno) Then
                    Record.SubItems(2) = !Bordxrefno
                End If
                If Len(!BordxClaimAmt) Then
                    Record.SubItems(3) = Format(!BordxClaimAmt, "#,##0.00")
                End If
                If Len(!BordxDate) Then
                   Record.SubItems(4) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If Len(!CASNumber) Then
                    Record.SubItems(5) = Trim(!CASNumber)
                End If
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(6) = Format(!CLMTotalApproved, "#,##0.00")
                End If
                If Len(!ClaimOpendate) Then
                    Record.SubItems(7) = Format(!ClaimOpendate, "dd/mm/yyyy")
                End If
                If Len(!BordxPAYRecDate) Then
                    Record.SubItems(8) = Format(!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                If Len(!ReceiptNoPAY) Then
                    Record.SubItems(9) = Trim(!ReceiptNoPAY)
                End If
                
                                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMTotalApproved) Then
                    TotalAmt = TotalAmt + rst2!CLMTotalApproved
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
    'cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'**********************End BordxTracking Listing *******************************

'**********************Start ComboListing *******************************

Private Sub PayorComboListing()
Const METHOD_NAME = "PayorComboListing"
On Error GoTo debugs
Dim PAY As New ADODB.Recordset
Dim Company As New ADODB.Recordset

    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
    PAY.Open "Select PAYCode, PAYCompanyName from PAY", MediConnCISDB, adOpenForwardOnly, adLockReadOnly ' adOpenKeyset, adLockOptimistic
    
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    Company.Open "Select GRPCoName from GRPCompany", MediConnCISDB, adOpenForwardOnly, adLockReadOnly 'adOpenKeyset, adLockOptimistic
    
    Do Until Company.EOF
        CboGrpCompany.AddItem Trim(Company!GrpCoName)
        Company.MoveNext
    Loop
    Company.Close
    Set Company = Nothing
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'**********************End ComboListing *******************************

'**********************Start ComboBox Change/KeyPress *******************************


Private Sub cboPayor_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "cboPayor_KeyPress"
On Error GoTo debugs
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayor.Text = ValidValue
                CboPayor.SelStart = 0
                CboPayor.SelLength = Len(ValidValue)
            End If
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cboPayor_Change()

If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
End Sub
'**********************End ComboBox Change/KeyPress *******************************

'**********************Start Sorted ListView *******************************
Private Sub lstBordxTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Const METHOD_NAME = "lstBordxTrackList_ColumnClick"
    On Error GoTo debugs
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(10)
    End Select
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'**********************End Sorted ListView *******************************
