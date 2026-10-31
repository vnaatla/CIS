VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMembershipTransfer 
   Caption         =   "Case Transfer"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4485
      Left            =   40
      TabIndex        =   6
      Top             =   240
      Width           =   12735
      Begin VB.Frame Frame2 
         Caption         =   "Current Member"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   16
         Top             =   680
         Width           =   5295
         Begin VB.TextBox txtcurrentdeptcode 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   40
            Top             =   2520
            Width           =   3855
         End
         Begin VB.TextBox txtcurrentCC 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   39
            Top             =   2160
            Width           =   3855
         End
         Begin VB.TextBox txtcurrentdeptname 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   38
            Top             =   1800
            Width           =   3855
         End
         Begin VB.TextBox txtcurrentsub 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   37
            Top             =   1440
            Width           =   3855
         End
         Begin VB.TextBox txtcurrentcompany 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   36
            Top             =   1080
            Width           =   3855
         End
         Begin VB.TextBox txtMBMNo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   18
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtCovID1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   17
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtMBMName1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Top             =   615
            Width           =   2175
         End
         Begin MSMask.MaskEdBox txtVisitDate 
            Height          =   315
            Left            =   4080
            TabIndex        =   20
            Top             =   615
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            Caption         =   "Dept Code"
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
            Left            =   120
            TabIndex        =   30
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Cost Center"
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
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Dept Name"
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
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   975
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
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Company Name"
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
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Member No."
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
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Cov ID"
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
            Left            =   3360
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Patient"
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
            Left            =   120
            TabIndex        =   22
            Top             =   675
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "Visit Date"
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
            Left            =   3360
            TabIndex        =   21
            Top             =   675
            Width           =   675
         End
      End
      Begin VB.Frame Frame5 
         Height          =   530
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   12495
         Begin VB.TextBox txtCaseNo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   0
            Top             =   160
            Width           =   2295
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "&Go"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3480
            TabIndex        =   13
            Top             =   160
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Case No."
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
            Left            =   120
            TabIndex        =   14
            Top             =   200
            Width           =   855
         End
      End
      Begin VB.Frame FrameTransfer 
         Caption         =   "Transfer Member"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   5430
         TabIndex        =   7
         Top             =   680
         Width           =   7215
         Begin VB.CommandButton cmbtnsearch 
            Caption         =   "Search"
            Height          =   375
            Left            =   5160
            TabIndex        =   51
            Top             =   540
            Width           =   1335
         End
         Begin VB.CheckBox Check5 
            Caption         =   "If Data Correct"
            Height          =   375
            Left            =   5760
            TabIndex        =   50
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            Caption         =   "If Data Correct"
            Height          =   375
            Left            =   5760
            TabIndex        =   49
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CheckBox Check3 
            Caption         =   "If Data Correct"
            Height          =   375
            Left            =   5760
            TabIndex        =   48
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "If Data Correct"
            Height          =   375
            Left            =   5760
            TabIndex        =   47
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "If Data Correct"
            Height          =   375
            Left            =   5760
            TabIndex        =   46
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   1440
            TabIndex        =   45
            Top             =   2400
            Width           =   4215
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1440
            TabIndex        =   44
            Top             =   2040
            Width           =   4215
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1440
            TabIndex        =   43
            Top             =   1680
            Width           =   4215
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1440
            TabIndex        =   42
            Top             =   1320
            Width           =   4215
         End
         Begin VB.TextBox txttracompanyname 
            Height          =   375
            Left            =   1440
            TabIndex        =   41
            Top             =   960
            Width           =   4215
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4800
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtMBMNo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Top             =   340
            Width           =   2295
         End
         Begin VB.TextBox txtCovID2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4000
            TabIndex        =   1
            Top             =   340
            Width           =   735
         End
         Begin VB.TextBox txtMBMName2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Top             =   620
            Width           =   3975
         End
         Begin VB.Label Label18 
            Caption         =   "Dept Code"
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
            Left            =   120
            TabIndex        =   35
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Cost Center"
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
            Left            =   120
            TabIndex        =   34
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Dept Name"
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
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label12 
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
            Left            =   120
            TabIndex        =   32
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Company Name"
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
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Patient"
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
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Cov ID"
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
            Left            =   3480
            TabIndex        =   10
            Top             =   340
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Transfer No."
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
            Left            =   120
            TabIndex        =   9
            Top             =   340
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   530
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   12495
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Transfer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10800
            TabIndex        =   4
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton CmdExit 
            Cancel          =   -1  'True
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11640
            TabIndex        =   5
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Case - Transfer of Membership "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmMembershipTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbtnsearch_Click()
Const METHOD_NAME = "cmbtnsearch_Click"
Dim rst2 As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2, prm3 As New ADODB.Parameter
On Error GoTo debugs
    Screen.MousePointer = vbHourglass
    txtMBMName2 = ""
    If txtMBMNo2 = "" Then
        MsgBox "Please enter Member Number! ", vbCritical
        txtMBMNo2.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf txtCovID2 = "" Then
        MsgBox "Please select Cover ID! ", vbCritical
        txtCovID2.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
    Set cmd1.ActiveConnection = medixmasterconnWS
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_GETcasetransfer"
    Set Prm1 = cmd1.CreateParameter("MemberNumber", adVarChar, adParamInput, 2000, Trim(txtMBMNo2))
    cmd1.Parameters.Append Prm1
    Set prm2 = cmd1.CreateParameter("CoverID", adVarChar, adParamInput, 1, Trim(txtCovID2))
    cmd1.Parameters.Append prm2
    Set prm3 = cmd1.CreateParameter("strAppendCase", adVarChar, adParamInput, 2000, 2)
    cmd1.Parameters.Append prm3
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd1.Execute
        Do Until rst2.EOF
            ''cboGroupCompany.AddItem Trim(rst2!GrpCoName)
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    End If
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()

Call ActionLogUpdate(12, "frmMembershipTransfer")

End Sub

Private Sub ListCases()
Const METHOD_NAME = "ListCases"
On Error GoTo debugs

Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset
    strsql = ""
    txtMBMName2 = ""
    txtMBMNo2 = ""
    txtCovID2 = ""
    strsql = " Select CASNumber, MBMCName, " & _
             " ReasonDesc, CoverID, CLMMCOProcFee, " & _
             " DateVisit, MBMCRELCode, CLMTotalActual,CLMBordxAmt, " & _
             " CASStatus, MBMName, MBMNumber, Claimability "
    strsql = strsql & " From VIEW_CISMBMTransfer where 0=0"
    strsql = strsql & " AND CASNumber = '" & RemStr(txtCaseNo) & "' "
            
    CAS.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If CAS.RecordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        txtCovID1 = CAS!coverid
        txtMBMNo1 = IIf(Len(CAS!MBMNumber), CAS!MBMNumber, "")
        If CAS!coverid = "00" Then
            
            If Len(CAS!MBMName) Then
                txtMBMName1 = Trim(CAS!MBMName)
            End If
        Else
            If Len(CAS!MBMCName) Then
                txtMBMName1 = Trim(CAS!MBMCName)
            End If
        End If
        
        If Len(Trim(CAS!DateVisit)) Then
            txtVisitDate = CAS!DateVisit
        End If
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Record not found!", vbCritical
        cmdSave.Enabled = False
        Screen.MousePointer = vbDefault
    End If
    CAS.Close
    Set CAS = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGo_Click()
Dim tmpCASNumber As String

    Screen.MousePointer = vbHourglass
        tmpCASNumber = txtCaseNo
        Call ClearForm(Me)
        txtCaseNo = tmpCASNumber
        If Trim(txtCaseNo) <> "" Then
            cmdSave.Enabled = False

          '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
          '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
          '  MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
                Call ListCases
        
           ' MediConnCISDB.Close
           ' Set MediConnCISDB = Nothing
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
           ' MediConnCISWS.Close
           ' Set MediConnCISWS = Nothing
            
        Else
            MsgBox "Please select Case Number ! ", vbCritical
        End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()

Const METHOD_NAME = "cmdSave_Click"
On Error GoTo debugs

Dim strsql As String

       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            
            strsql = " Update CISCAS set " & _
                     " MBMNumber = '" & txtMBMNo2 & "', CoverID = '" & txtCovID2 & "'" & _
                     " WHERE CASNumber = '" & Trim(txtCaseNo) & "'"
            MediConnCISDB.Execute strsql
            'Call ListCases
            cmdSave.Enabled = False
        
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'MediConnCISWS.Close
        'Set MediConnCISWS = Nothing
        
        MsgBox "Record Updated!", vbInformation
        
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs


Dim strsql As String
Dim MBM As New ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    txtMBMName2 = ""
    If txtMBMNo2 = "" Then
        MsgBox "Please enter Member Number! ", vbCritical
        txtMBMNo2.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf txtCovID2 = "" Then
        MsgBox "Please select Cover ID! ", vbCritical
        txtCovID2.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        If Trim(txtCovID2) = "00" Then
            strsql = " Select MBMName From MBM where MBMNumber = '" & RemStr(txtMBMNo2) & "'"
            MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        Else
            strsql = " Select MBMCName As MBMName From MBMCoveredPersons where MBMCNumber = '" & RemStr(txtMBMNo2) & "' And MBMCCoverID = '" & Trim(txtCovID2) & "'"
            MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        End If
        
        If MBM.RecordCount <> 0 Then
            If Len(MBM!MBMName) Then
                txtMBMName2 = MBM!MBMName
            End If
            cmdSave.Enabled = True
        Else
            MsgBox "Record not found!", vbCritical
            cmdSave.Enabled = False
        End If
        MBM.Close
        Set MBM = Nothing
    End If
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub



Private Sub txtCaseNo_Change()
    cmdSave.Enabled = False
End Sub

Private Sub txtCovID2_Change()
    cmdSave.Enabled = False
End Sub


Private Sub txtMBMNo2_Change()
    cmdSave.Enabled = False
End Sub
