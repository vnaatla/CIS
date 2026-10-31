VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimsdata 
   Caption         =   "Export Claims Data"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   240
   ClientWidth     =   14028
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   14028
   WindowState     =   2  'Maximized
   Begin VB.Frame frmmain 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10215
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   4080
         TabIndex        =   19
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdexport 
         Caption         =   "&Export"
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   2400
         Width           =   1095
      End
      Begin VB.ComboBox cbosubmited 
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
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   1570
      End
      Begin VB.TextBox TxtWSNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TxtWSNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1320
         Width           =   1605
      End
      Begin VB.ComboBox CboGrpCom 
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
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   885
         Width           =   4220
      End
      Begin VB.ComboBox CboPayor 
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
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   4220
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lsTrackList 
         Height          =   2985
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   9735
         _ExtentX        =   17166
         _ExtentY        =   5271
         View            =   3
         LabelEdit       =   1
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bordx No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Claim Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Visit Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sent Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Patient Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cash Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label25 
         Caption         =   "Claim No"
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
         TabIndex        =   14
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "To"
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
         Left            =   3240
         TabIndex        =   13
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label txtStatus 
         Caption         =   "Submited"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label49 
         Caption         =   "Grp Company"
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
         TabIndex        =   8
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
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
         Top             =   705
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Index           =   1
         Left            =   3000
         TabIndex        =   4
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx Date"
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Date of Visit"
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "Generate Claims Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmClaimsdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboGrpCom_Change()
    If CboGrpCom.Text = "" Then
        If InStr(CboGrpCom.Text, "null") Then
            CboGrpCom.Enabled = False
        End If
    End If
End Sub

Private Sub cmdclear_Click()
    txtBordxDateFr = "__/__/____"
    txtBordxDateTo = "__/__/____"
    CboGrpCom = ""
    CboPayor = ""
    cbosubmited = ""
    TxtWSNo1 = ""
    TxtWSNo2 = ""
    
End Sub

Private Sub cmdexit_Click()
 Unload Me
End Sub

Private Sub cmdsearch_Click()
    Call SearchRecord
End Sub
Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs
Dim strAppend As String
strAppend = ""

    If Len(Trim(TxtWSNo1)) Then
        strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtWSNo1) & "'"
    End If
    
    If Len(Trim(TxtWSNo2)) Then
        strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtWSNo2) & "'"
    End If
    
    If Len(Trim(CboGrpCom)) Then
       strAppend = strAppend + " AND MBMGroupCompany = '" & Trim(RemStr(CboGrpCom)) & "'"
    End If
    
     If Len(Trim(CboPayor)) Then
       strAppend = strAppend + " AND MBMGroupCompany = '" & Mid(CboPayor, 1, 2) & "'"
    End If
    
    If Len(Trim(txtBordxDateFr)) > 0 Then
        strAppend = strAppend + " AND BordxRefNo >= '" & Trim(txtBordxDateFr) & "'"
    End If
    
    If Len(Trim(txtBordxDateTo)) > 0 Then
        strAppend = strAppend + " AND BordxRefNo <= '" & Trim(txtBordxDateTo) & "'"
    End If
    
    Call TrackingListing(strAppend)

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
Private Sub TrackingListing(Optional strAppend As String)
Const METHOD_NAME = "TrackingListing"
On Error GoTo debugs


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
    Call ActionLogUpdate(12, "frmClaimsdata")
    Call ComboListing
End Sub
Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs
Dim PAY As New ADODB.Recordset
Dim monthstart
Dim yearstart
'Group company
        PAY.Open "Select DISTINCT GRPCoName from GRPCompany ", MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do Until PAY.EOF
            CboGrpCom.AddItem Trim(PAY!GrpCoName)
            PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
    'Submit
    With cbosubmited
        .AddItem "YES"
        .AddItem "NO"
    End With
    'Pay code
    PAY.Open "select PAYCode, PAYCompanyName from PAY ", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do While Not PAY.EOF
        CboPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing

    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub



