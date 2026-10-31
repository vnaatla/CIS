VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxControl 
   Caption         =   "Bordx Control"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   410
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Bordx Control"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   4440
         TabIndex        =   50
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   8640
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7800
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6960
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9600
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstBordxNonsihat 
      Height          =   4935
      Left            =   0
      TabIndex        =   44
      Top             =   2280
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8705
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Bordx Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "No of Claims"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bill Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Co-Pay"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "App Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "FFS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Bordx Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Non GP"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DTGrp Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "DRGrp Company (Doc)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "No of Days (DRGrp Company - DTGrp Company)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "DFGrp Company (Money)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "No of Days (DTGrp Company - DPayment)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Subsidiary"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Cash Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   0
      TabIndex        =   27
      Top             =   360
      Width           =   11895
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "FrmBordxControl.frx":0000
         Left            =   1320
         List            =   "FrmBordxControl.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1500
         Width           =   3015
      End
      Begin VB.ComboBox Cbocashtype 
         Height          =   315
         ItemData        =   "FrmBordxControl.frx":002B
         Left            =   7680
         List            =   "FrmBordxControl.frx":002D
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cboSubsidiary 
         Height          =   315
         ItemData        =   "FrmBordxControl.frx":002F
         Left            =   1320
         List            =   "FrmBordxControl.frx":0031
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Frame Frame10 
         Caption         =   "From Payor (Money)"
         Height          =   1275
         Left            =   10030
         TabIndex        =   37
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton ChkPaid3 
            Caption         =   "Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton ChkPartial3 
            Caption         =   "Partial Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton ChkUnpaid3 
            Caption         =   "Unpaid"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   1455
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame5"
            Height          =   615
            Left            =   1800
            TabIndex        =   38
            Top             =   120
            Width           =   15
         End
         Begin VB.OptionButton optBoth4 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "From MNRB (Money)"
         Height          =   1275
         Left            =   10080
         TabIndex        =   31
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton optBoth3 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame5"
            Height          =   615
            Left            =   1800
            TabIndex        =   35
            Top             =   120
            Width           =   15
         End
         Begin VB.OptionButton ChkUnpaid2 
            Caption         =   "Unpaid"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton ChkPartial2 
            Caption         =   "Partial Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton ChkPaid2 
            Caption         =   "Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "From Payor (Doc)"
         Height          =   1275
         Left            =   8340
         TabIndex        =   29
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton ChkPaid1 
            Caption         =   "Received"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton ChkPartial1 
            Caption         =   "Not Yet Received"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame5"
            Height          =   615
            Left            =   1800
            TabIndex        =   30
            Top             =   120
            Width           =   15
         End
         Begin VB.OptionButton optBoth1 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "To Payor"
         Height          =   1275
         Left            =   6410
         TabIndex        =   28
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton ChkSent 
            Caption         =   "Sent"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton ChkNotSent 
            Caption         =   "Not Yet Sent"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optBoth 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox CboGrpCompany 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   950
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   950
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEntityType 
         Caption         =   "Entity Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   52
         ToolTipText     =   "Date Pass On"
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cash Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   47
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Subsidiary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No from"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   680
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   950
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   1030
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmBordxControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intExit As Integer
Private m_blnDirection(1 To 17) As Boolean

Private Sub cboGRPCompany_LostFocus()
Const METHOD_NAME = "cboGRPCompany_LostFocus"
On Error GoTo debugs
Call BindSubsidary
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
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

Private Sub CmdPrintFile_Click()
Const METHOD_NAME = "CmdPrintFile_Click"
On Error GoTo debugs
    Screen.MousePointer = vbHourglass
        
        Call ExportToExcelBordxControlNonSihat(lstBordxNonsihat)
    
    Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs
Call ActionLogUpdate(12, "FrmBordxControl")
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    cboEntityType.ListIndex = 0
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        Call PayorComboListing
       ' Call BindSubsidary
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub BindSubsidary()
Const METHOD_NAME = "BindSubsidary"
On Error GoTo debugs
  Dim rst3 As New ADODB.Recordset
  Dim sql2 As String
  
        sql2 = ""
        sql2 = "Exec SearchSubsidiary '" & Trim(Replace(CboGrpCompany, "'", "''")) & "'"
        rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
  
        cboSubsidiary.Clear
        Do Until rst3.EOF
            With rst3
                cboSubsidiary.AddItem IIf(Len(Trim(!MBMGRPSubsidiary)), Trim(!MBMGRPSubsidiary), "")
                rst3.MoveNext
            End With
        Loop
        rst3.Close
        Set rst3 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'************************** Start Command Button **************************
Private Sub cmdSearch_Click()
    Const METHOD_NAME = "cmdSearch_Click"
    On Error GoTo debugs
    LblTotalAmt = ""
    lblCurrent = ""
    lblTotal = ""
    
    CmdClear.Enabled = False
    CmdPrintFile.Enabled = False
    CmdExit.Enabled = False
    
    Screen.MousePointer = vbHourglass
        Call SearchRecord
    Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdStop_Click()
Const METHOD_NAME = "CmdStop_Click"
On Error GoTo debugs
    intExit = 1
    CmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CboGrpCompany.Clear
    Call PayorComboListing
    Call BindSubsidary
    Frame9.Enabled = True
    Frame3.Enabled = True
    lstBordxNonsihat.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
End Sub
'************************** End Command Button **************************

'************************** Start Search Record **************************
Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs
Dim strsql As String
Dim StrAppend As String
Dim Sql As String
Dim CAS As New ADODB.Recordset
Dim getSel As Boolean
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim GrpCompanyName As String
    
    StrAppend = ""
    GrpCompanyName = ""
    
    If Len(Trim(CboGrpCompany)) Then
         GrpCompanyName = Replace(CboGrpCompany, "'", "''")
         StrAppend = StrAppend + " AND Bordx.GroupCompany = '" & Trim(GrpCompanyName) & "'"
    End If
       
    If Len(Trim(TxtBordx1)) Then
        StrAppend = StrAppend + " AND Bordx.BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
        StrAppend = StrAppend + " AND Bordx.BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
    End If
    
    If IsDate(txtBordxDate1) = True Then
        StrAppend = StrAppend + " And Bordx.BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtBordxDate2) = True Then
        StrAppend = StrAppend + " And Bordx.BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkSent.Value = True Then
        StrAppend = StrAppend + " AND BordxPAYToStatus = '1'"
    End If
    
    If ChkNotSent.Value = True Then
        StrAppend = StrAppend + " AND BordxPAYToStatus = '0'"
    End If
    
    If ChkPaid1.Value = True Then
        StrAppend = StrAppend + " AND BordxPAYDOCFrStatus = '1'"
    End If

    If ChkPartial1.Value = True Then
        StrAppend = StrAppend + " AND BordxPAYDOCFrStatus = '0'"
    End If

    If ChkPaid3.Value = True Then
        StrAppend = StrAppend + " AND BordxPAYFrStatus = 'F'"
    End If

    If ChkPartial3.Value = True Then
        StrAppend = StrAppend + " AND BordxPAYFrStatus = 'P'"
    End If

    If ChkUnpaid3.Value = True Then
        StrAppend = StrAppend + " AND (BordxPAYFrStatus = '' or BordxPAYFrStatus is null)"
    End If
    
    If Len(Trim(cboSubsidiary)) Then
        StrAppend = StrAppend + " AND BordxSubsidary = '" & Trim(cboSubsidiary) & "'"
    End If
    
    If Mid(Cbocashtype, 1, 1) = "C" Then
        StrAppend = StrAppend & " AND INVPayTo = 'C'"
    ElseIf Mid(Cbocashtype, 1, 1) = "M" Then
        StrAppend = StrAppend & " AND INVPayTo = 'M' "
    End If
    
    If cboEntityType.ListIndex = 0 Then
        StrAppend = StrAppend + " AND Substring(CISCas.CasNumber, 1, 2) <>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        StrAppend = StrAppend + " AND Substring(CISCas.CasNumber, 1, 2) ='SJ' "
    End If
        
    Call NonSihatBordxControlListing(StrAppend)
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************************** End Search Record **************************

'************************** Start Non Sihat Bordx Control Listing **************************
Private Sub NonSihatBordxControlListing(Optional StrAppend As String)
Const METHOD_NAME = "NonSihatBordxControlListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim rstCount  As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql3 As String
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
        
    lstBordxNonsihat.ListItems.Clear
            
   
    
    Sql = Sql + StrAppend
            
    Set cmd.ActiveConnection = MediConnCISWS
    If Len(Trim(cboSubsidiary)) Then
         strAppendCase = "16"
        cmd.CommandText = "ClaimBordxModule"
    Else
         strAppendCase = "17"
        cmd.CommandText = "ClaimBordxModule1"
    End If
    
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 2000
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, Sql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    'Sql = " Select GroupCompany, BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt, " & _
          " BordxPAYSentDate, BordxPAYRecDocDate, BordxPAYRecDate, CLMNonGpFee, " & _
          " DateDiff(d, BordxPAYSentDate, BordxPAYRecDocDate) AS NoofDay, " & _
          " DateDiff(d, BordxPAYSentDate, BordxPAYRecDate) AS NoofDay2 " & _
          " From VIEW_CISBordx_ControlN where 0 = 0 "
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
                Set Record = lstBordxNonsihat.ListItems.Add(, , !Bordxrefno)
                If Len(!BordxDate) Then
                    Record.SubItems(1) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                
                If Len(!BordxNoClaims) Then
                   Record.SubItems(2) = !BordxNoClaims
                End If
                
                Record.SubItems(3) = IIf(IsNumeric(!BillAmt), Format(!BillAmt, "#,##0.00"), "")
                Record.SubItems(4) = IIf(IsNumeric(!CLMCoPayCLNColl), Format(!CLMCoPayCLNColl, "#,##0.00"), "")
                Record.SubItems(5) = IIf(IsNumeric(!CLMFinalApprove), Format(!CLMFinalApprove, "#,##0.00"), "")
                Record.SubItems(6) = IIf(IsNumeric(!CLMMCOProcFee), Format(!CLMMCOProcFee, "#,##0.00"), "")
                Record.SubItems(7) = IIf(IsNumeric(!CLMBordxAmt), Format(!CLMBordxAmt, "#,##0.00"), "")
                Record.SubItems(8) = IIf(IsNumeric(!CLMNonGpFee), Format(!CLMNonGpFee, "#,##0.00"), "")
               
                If Len(!BordxPAYSentDate) Then
                    Record.SubItems(9) = Format(!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                
                If Len(!BordxPAYRecDocDate) Then
                    Record.SubItems(10) = Format(!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                
                If !NoofDay = "0" Then
                    Record.SubItems(11) = "1"
                ElseIf Len(!NoofDay) Then
                    Record.SubItems(11) = !NoofDay
                End If
                
                If Len(!BordxPAYRecDate) Then
                    Record.SubItems(12) = Format(!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                
                If !NoofDay2 = "0" Then
                    Record.SubItems(13) = "1"
                ElseIf Len(!NoofDay2) Then
                    Record.SubItems(13) = !NoofDay2
                End If
                
                If Len(!MBMGroupCompany) Then
                    Record.SubItems(14) = !MBMGroupCompany
                End If
                
                If Len(Trim(cboSubsidiary)) Then
                    If Len(!Subsidiary) Then
                        Record.SubItems(15) = IIf(Len(Trim(!Subsidiary)), Trim(!Subsidiary), "")
                    End If
                Else
                    Record.SubItems(15) = IIf(Len(Trim(!Subsidiary)), Trim(!Subsidiary), "")
                End If
                If Len(!INVPayTo) Then
                    Record.SubItems(16) = !INVPayTo
                End If
               
                Current = Current + 1
                lblCurrent = Current
                If IsNumeric(Record.SubItems(6)) Then
                    TotalAmt = CDbl(IIf(Len(TotalAmt), TotalAmt, 0)) + CDbl(Record.SubItems(6))
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
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'************************** End Non sihat Bordx Control Listing **************************

'************************** Start ComboListing **************************

Private Sub PayorComboListing()
Const METHOD_NAME = "PayorComboListing"
On Error GoTo debugs
Dim GrpCompany As New ADODB.Recordset
    GrpCompany.Open "Select GRPCoName from GRPCompany", MediConnCISDB, adOpenForwardOnly, adLockReadOnly 'adOpenKeyset, adLockOptimistic
    
    Do Until GrpCompany.EOF
        CboGrpCompany.AddItem Trim(GrpCompany!GrpCoName)
        GrpCompany.MoveNext
    Loop
    GrpCompany.Close
    Set GrpCompany = Nothing
    With Cbocashtype
        .AddItem "C - Cashless"
        .AddItem "M - Member"
    End With

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'************************** End ComboListing **************************

Private Sub CboGrpCompany_Click()
Const METHOD_NAME = "CboGrpCompany_Click"
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

Private Sub CboGrpCompany_Change()
Const METHOD_NAME = "CboGrpCompany_Change"
On Error GoTo debugs

    If CboGrpCompany.Text = "" Then
        If InStr(CboGrpCompany.Text, "null") Then
           CboGrpCompany.Enabled = False
        End If
    End If
    
    Call BindSubsidary
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstBordxNonsihat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Const METHOD_NAME = "lstBordxNonsihat_ColumnClick"
    On Error GoTo debugs
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
        
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
    Case 3
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    End Select
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub TxtBordxDate1_LostFocus()
    Const METHOD_NAME = "TxtBordxDate1_LostFocus"
    On Error GoTo debugs
    If IsDate(txtBordxDate1) Then
    Else
        If txtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate1.SetFocus
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub TxtBordxDate2_LostFocus()
Const METHOD_NAME = "TxtBordxDate2_LostFocus"
On Error GoTo debugs
    If IsDate(txtBordxDate2) Then
    Else
        If txtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate2.SetFocus
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
