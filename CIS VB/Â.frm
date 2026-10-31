VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimRpt 
   Caption         =   "Claim Report"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12270
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CboSelection 
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2160
   End
   Begin VB.ComboBox CboReportType 
      Height          =   315
      ItemData        =   "Â.frx":0000
      Left            =   1320
      List            =   "Â.frx":004C
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1920
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   226
      Top             =   0
      Width           =   11925
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Claims Reports"
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
         TabIndex        =   227
         Top             =   0
         Width           =   3285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   220
      Top             =   8160
      Width           =   11775
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   232
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtDay 
         Height          =   285
         Left            =   1440
         TabIndex        =   223
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   85
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   90
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   91
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   89
         Top             =   160
         Width           =   975
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   222
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   221
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7080
         TabIndex        =   87
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   88
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6240
         TabIndex        =   86
         Top             =   160
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   255
         Left            =   3600
         TabIndex        =   224
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   32
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   225
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Selection Criteria"
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   12165
      Begin VB.CheckBox ChkScan 
         Caption         =   "Append"
         Height          =   255
         Left            =   5280
         TabIndex        =   242
         Top             =   5350
         Width           =   1455
      End
      Begin VB.OptionButton OPTBordxNo 
         Caption         =   "Option1"
         Height          =   255
         Left            =   8280
         TabIndex        =   239
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton OPTClaimNo 
         Caption         =   "Option1"
         Height          =   255
         Left            =   8280
         TabIndex        =   238
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox txtImportFilenameSUN 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   237
         Top             =   5280
         Width           =   2265
      End
      Begin VB.CommandButton cmdFileLocationSUN 
         Caption         =   " >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   236
         Top             =   5280
         Width           =   435
      End
      Begin VB.CommandButton cmdUploadDirect 
         Caption         =   "Upload Excel"
         Height          =   375
         Left            =   3840
         TabIndex        =   235
         Top             =   5280
         Width           =   1335
      End
      Begin VB.ComboBox CboIEinv 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8520
         Sorted          =   -1  'True
         TabIndex        =   234
         Top             =   5040
         Width           =   840
      End
      Begin VB.CheckBox ChkGrpCom2 
         Height          =   255
         Left            =   6960
         TabIndex        =   231
         Top             =   500
         Width           =   255
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   285
         Left            =   8520
         MaxLength       =   60
         TabIndex        =   53
         Top             =   195
         Width           =   2880
      End
      Begin VB.CheckBox ChkAmt 
         Height          =   255
         Left            =   11520
         TabIndex        =   134
         Top             =   4800
         Width           =   255
      End
      Begin VB.CheckBox ChkDOW 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   6720
         TabIndex        =   133
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox ChkICNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   11520
         TabIndex        =   132
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxDate 
         Height          =   255
         Left            =   11520
         TabIndex        =   131
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox ChkPolYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   6720
         TabIndex        =   130
         Top             =   4200
         Width           =   255
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "N - Non-Sihat"
         Top             =   200
         Width           =   1320
      End
      Begin VB.CheckBox ChkDOA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   6720
         TabIndex        =   129
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox chkDOR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   6720
         TabIndex        =   128
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox ChkJoinDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   6720
         TabIndex        =   127
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox ChkDOB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6720
         TabIndex        =   126
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   6720
         TabIndex        =   125
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   6720
         TabIndex        =   124
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2520
         TabIndex        =   123
         Top             =   4950
         Width           =   255
      End
      Begin VB.CheckBox ChkHOS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   6720
         TabIndex        =   122
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox ChkFirst 
         Caption         =   "Check3"
         Height          =   255
         Left            =   6720
         TabIndex        =   121
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox ChkEffDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   6720
         TabIndex        =   120
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox ChkBordx 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   11520
         TabIndex        =   119
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox ChkPlan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   6720
         TabIndex        =   118
         Top             =   4920
         Width           =   375
      End
      Begin VB.CheckBox ChkClaimability 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   117
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox ChkRenewal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   116
         Top             =   4110
         Width           =   255
      End
      Begin VB.CheckBox ChkDataStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   115
         Top             =   2145
         Width           =   255
      End
      Begin VB.CheckBox ChkRelation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   114
         Top             =   1845
         Width           =   255
      End
      Begin VB.CheckBox ChkTakeOver 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   113
         Top             =   4395
         Width           =   255
      End
      Begin VB.CheckBox ChkAgentCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   112
         Top             =   4665
         Width           =   255
      End
      Begin VB.CheckBox ChkHLT 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2520
         TabIndex        =   111
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2520
         TabIndex        =   110
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkRace 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2520
         TabIndex        =   109
         Top             =   1305
         Width           =   255
      End
      Begin VB.CheckBox ChkSex 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2520
         TabIndex        =   108
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox ChkState 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   107
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox ChkExpired 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   106
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox ChkPolStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2520
         TabIndex        =   105
         Top             =   2685
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2520
         TabIndex        =   104
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox ChkPreMode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2520
         TabIndex        =   103
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox ChkMBMIC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   11520
         TabIndex        =   102
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox ChkMCDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   6720
         TabIndex        =   101
         Top             =   4440
         Width           =   255
      End
      Begin VB.CheckBox ChkSubmission 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   11520
         TabIndex        =   100
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox ChkInvoice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   11520
         TabIndex        =   99
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox ChkVisitDay 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   6720
         TabIndex        =   98
         Top             =   4680
         Width           =   255
      End
      Begin VB.CheckBox chkCLN 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6960
         TabIndex        =   97
         Top             =   1080
         Width           =   195
      End
      Begin VB.CheckBox chkPay 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6960
         TabIndex        =   96
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ChkInvoiceDate 
         Height          =   255
         Left            =   11520
         TabIndex        =   95
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox chkExpDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   6720
         TabIndex        =   94
         Top             =   3135
         Width           =   255
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "Check1"
         Height          =   255
         Left            =   6720
         TabIndex        =   93
         Top             =   790
         Width           =   255
      End
      Begin VB.CheckBox chkDept2 
         Caption         =   "Check1"
         Height          =   255
         Left            =   6960
         TabIndex        =   92
         Top             =   790
         Width           =   255
      End
      Begin VB.ComboBox CboInsType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1320
      End
      Begin VB.ComboBox CboMemType 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1320
      End
      Begin VB.ComboBox cboSuppType 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1000
         Width           =   1320
      End
      Begin VB.ComboBox CboRace 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1260
         Width           =   1320
      End
      Begin VB.ComboBox CboSex 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1545
         Width           =   1320
      End
      Begin VB.ComboBox CboRelation 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1830
         Width           =   1320
      End
      Begin VB.ComboBox cboDataStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   2115
         Width           =   1320
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2400
         Width           =   1320
      End
      Begin VB.ComboBox CboStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   2650
         Width           =   1320
      End
      Begin VB.ComboBox CboCaseStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   2940
         Width           =   1320
      End
      Begin VB.ComboBox CboExpired 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   3230
         Width           =   1320
      End
      Begin VB.ComboBox CboPreMode 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   3500
         Width           =   1320
      End
      Begin VB.ComboBox CboState 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   3780
         Width           =   1320
      End
      Begin VB.ComboBox cboRenewalInd 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   4070
         Width           =   1320
      End
      Begin VB.ComboBox cboTakeOver 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   4350
         Width           =   1320
      End
      Begin VB.TextBox txtAgentCode 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   19
         Top             =   4640
         Width           =   1320
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   20
         Top             =   4890
         Width           =   1320
      End
      Begin VB.ComboBox cboPayorCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   200
         Width           =   2640
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   2640
      End
      Begin VB.ComboBox cboDeptSelection 
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   760
         Width           =   2655
      End
      Begin VB.ComboBox cboCLNName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   1050
         Width           =   2640
      End
      Begin VB.ComboBox cboCLNGrp 
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   1340
         Width           =   2640
      End
      Begin VB.TextBox TxtBranch 
         Height          =   285
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1620
         Width           =   2640
      End
      Begin MSMask.MaskEdBox TxtDOV1 
         Height          =   315
         Left            =   3960
         TabIndex        =   27
         Top             =   1860
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR1 
         Height          =   315
         Left            =   3960
         TabIndex        =   29
         Top             =   2145
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOW1 
         Height          =   315
         Left            =   3960
         TabIndex        =   31
         Top             =   2400
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtTreatment1 
         Height          =   285
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   33
         Top             =   2640
         Width           =   1080
      End
      Begin MSMask.MaskEdBox TxtDOV2 
         Height          =   315
         Left            =   5520
         TabIndex        =   28
         Top             =   1860
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   315
         Left            =   5520
         TabIndex        =   30
         Top             =   2145
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOW2 
         Height          =   315
         Left            =   5520
         TabIndex        =   32
         Top             =   2400
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtTreatment2 
         Height          =   285
         Left            =   5520
         MaxLength       =   3
         TabIndex        =   34
         Top             =   2640
         Width           =   1080
      End
      Begin MSMask.MaskEdBox TxtPayEffDate1 
         Height          =   315
         Left            =   3960
         TabIndex        =   35
         Top             =   2880
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   5520
         TabIndex        =   36
         Top             =   2880
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayExpDate1 
         Height          =   315
         Left            =   3960
         TabIndex        =   37
         Top             =   3120
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin1 
         Height          =   315
         Left            =   3960
         TabIndex        =   39
         Top             =   3360
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB1 
         Height          =   315
         Left            =   3960
         TabIndex        =   41
         Top             =   3600
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtAgeband1 
         Height          =   285
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   43
         Top             =   3885
         Width           =   1080
      End
      Begin VB.TextBox txtPOLYear1 
         Height          =   285
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   45
         Top             =   4125
         Width           =   1080
      End
      Begin VB.TextBox TxtMCDay1 
         Height          =   285
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   47
         Top             =   4365
         Width           =   1080
      End
      Begin VB.TextBox TxtVisitDay1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   49
         Top             =   4620
         Width           =   1080
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   285
         Left            =   3960
         MaxLength       =   4
         TabIndex        =   51
         Top             =   4860
         Width           =   1080
      End
      Begin MSMask.MaskEdBox TxtPayExpDate2 
         Height          =   315
         Left            =   5520
         TabIndex        =   38
         Top             =   3120
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin2 
         Height          =   315
         Left            =   5520
         TabIndex        =   40
         Top             =   3360
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB2 
         Height          =   315
         Left            =   5520
         TabIndex        =   42
         Top             =   3600
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtAgeband2 
         Height          =   285
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   44
         Top             =   3885
         Width           =   1080
      End
      Begin VB.TextBox txtPOLYear2 
         Height          =   285
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   46
         Top             =   4125
         Width           =   1080
      End
      Begin VB.TextBox TxtMCDay2 
         Height          =   285
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   48
         Top             =   4365
         Width           =   1080
      End
      Begin VB.TextBox TxtVisitDay2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   50
         Top             =   4620
         Width           =   1080
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   285
         Left            =   5520
         MaxLength       =   4
         TabIndex        =   52
         Top             =   4860
         Width           =   1080
      End
      Begin VB.TextBox TxtICNo1 
         Height          =   285
         Left            =   8520
         MaxLength       =   20
         TabIndex        =   54
         Top             =   435
         Width           =   1200
      End
      Begin VB.TextBox TxtICNo2 
         Height          =   285
         Left            =   10200
         MaxLength       =   20
         TabIndex        =   55
         Top             =   435
         Width           =   1200
      End
      Begin VB.TextBox TxtMBMIc1 
         Height          =   285
         Left            =   8520
         MaxLength       =   20
         TabIndex        =   56
         Top             =   690
         Width           =   1200
      End
      Begin VB.TextBox TxtMBMIc2 
         Height          =   285
         Left            =   10200
         MaxLength       =   20
         TabIndex        =   57
         Top             =   690
         Width           =   1200
      End
      Begin VB.TextBox txtActualAmt1 
         Height          =   285
         Left            =   8520
         TabIndex        =   58
         Top             =   945
         Width           =   1200
      End
      Begin VB.TextBox txtActualAmt2 
         Height          =   285
         Left            =   10200
         TabIndex        =   59
         Top             =   945
         Width           =   1200
      End
      Begin VB.TextBox txtAppAmt1 
         Height          =   285
         Left            =   8520
         TabIndex        =   60
         Top             =   1200
         Width           =   1200
      End
      Begin VB.TextBox txtAppAmt2 
         Height          =   285
         Left            =   10200
         TabIndex        =   61
         Top             =   1200
         Width           =   1200
      End
      Begin VB.TextBox TxtInvoice1 
         Height          =   285
         Left            =   8520
         MaxLength       =   20
         TabIndex        =   62
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox TxtInvoice2 
         Height          =   285
         Left            =   10200
         MaxLength       =   20
         TabIndex        =   63
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox TxtSubmission1 
         Height          =   285
         Left            =   8520
         MaxLength       =   20
         TabIndex        =   64
         Top             =   1695
         Width           =   1200
      End
      Begin VB.TextBox TxtSubmission2 
         Height          =   285
         Left            =   10200
         MaxLength       =   20
         TabIndex        =   65
         Top             =   1680
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtInvDate1 
         Height          =   315
         Left            =   8520
         TabIndex        =   66
         Top             =   1920
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInvDate2 
         Height          =   315
         Left            =   10200
         TabIndex        =   67
         Top             =   1920
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   285
         Left            =   8520
         MaxLength       =   15
         TabIndex        =   68
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   285
         Left            =   10200
         MaxLength       =   15
         TabIndex        =   69
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   8520
         MaxLength       =   15
         TabIndex        =   70
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   285
         Left            =   10200
         MaxLength       =   15
         TabIndex        =   71
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   72
         Top             =   2640
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   10200
         MaxLength       =   10
         TabIndex        =   73
         Top             =   2640
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   8520
         TabIndex        =   74
         Top             =   2880
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordxDate2 
         Height          =   315
         Left            =   10200
         TabIndex        =   75
         Top             =   2880
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtOSUtilized1 
         Height          =   285
         Left            =   8520
         MaxLength       =   3
         TabIndex        =   76
         Top             =   3165
         Width           =   1200
      End
      Begin VB.TextBox TxtOSUtilized2 
         Height          =   285
         Left            =   10200
         MaxLength       =   3
         TabIndex        =   77
         Top             =   3165
         Width           =   1200
      End
      Begin VB.TextBox TxtOPBatchNo1 
         Height          =   285
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   78
         Top             =   3420
         Width           =   1200
      End
      Begin VB.TextBox TxtOPBatchNo2 
         Height          =   285
         Left            =   10200
         MaxLength       =   10
         TabIndex        =   79
         Top             =   3420
         Width           =   1200
      End
      Begin VB.ComboBox CboClinicType 
         Height          =   315
         Left            =   8520
         Sorted          =   -1  'True
         TabIndex        =   80
         Top             =   3675
         Width           =   2880
      End
      Begin VB.ComboBox cboPanelStatus 
         Height          =   315
         Left            =   8520
         Sorted          =   -1  'True
         TabIndex        =   81
         Top             =   3945
         Width           =   2880
      End
      Begin VB.ComboBox CboClmPending 
         Height          =   315
         Left            =   8520
         Sorted          =   -1  'True
         TabIndex        =   82
         Top             =   4200
         Width           =   2880
      End
      Begin VB.ComboBox CboPayType 
         Height          =   315
         Left            =   8520
         Sorted          =   -1  'True
         TabIndex        =   83
         Top             =   4485
         Width           =   2880
      End
      Begin VB.TextBox TxtYear 
         Height          =   285
         Left            =   8520
         MaxLength       =   4
         TabIndex        =   84
         Top             =   4770
         Width           =   1080
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11160
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label30 
         Caption         =   "File Name"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   241
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   144
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label90 
         Caption         =   "IE Invoice"
         Height          =   255
         Left            =   7440
         TabIndex        =   233
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label Label42 
         Caption         =   "DOB"
         Height          =   255
         Left            =   2880
         TabIndex        =   219
         ToolTipText     =   "Date of Birth"
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Branch"
         Height          =   255
         Left            =   2880
         TabIndex        =   218
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label76 
         Caption         =   "DOW"
         Height          =   255
         Left            =   2880
         TabIndex        =   217
         ToolTipText     =   "Worksheet Date"
         Top             =   2460
         Width           =   855
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   216
         Top             =   2460
         Width           =   255
      End
      Begin VB.Label Label74 
         Caption         =   "Age Band"
         Height          =   255
         Left            =   2880
         TabIndex        =   215
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   214
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label84 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   213
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label85 
         Caption         =   "Patient I/C"
         Height          =   255
         Left            =   7440
         TabIndex        =   212
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   211
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label70 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   7440
         TabIndex        =   210
         Top             =   2985
         Width           =   975
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   209
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label66 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   7440
         TabIndex        =   208
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Member No"
         Height          =   255
         Left            =   7440
         TabIndex        =   207
         Top             =   2180
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   120
         TabIndex        =   206
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   205
         Top             =   4950
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "State"
         Height          =   255
         Left            =   120
         TabIndex        =   204
         Top             =   3870
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Prem. Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   203
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Exp. Status"
         Height          =   255
         Left            =   120
         TabIndex        =   202
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Relationship"
         Height          =   255
         Left            =   120
         TabIndex        =   200
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Race"
         Height          =   255
         Left            =   120
         TabIndex        =   199
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   120
         TabIndex        =   198
         Top             =   830
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Sex"
         Height          =   255
         Left            =   120
         TabIndex        =   197
         Top             =   1590
         Width           =   855
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   196
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label47 
         Caption         =   "DOV"
         Height          =   255
         Left            =   2880
         TabIndex        =   195
         ToolTipText     =   "Date of Visit"
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   194
         Top             =   2205
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   193
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   192
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "DOR"
         Height          =   255
         Left            =   2880
         TabIndex        =   191
         ToolTipText     =   "Date of Register"
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   190
         ToolTipText     =   "Payor Expiry Date"
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   189
         ToolTipText     =   "Payor Effective Date"
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label73 
         Caption         =   "HLT Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   188
         Top             =   250
         Width           =   705
      End
      Begin VB.Label Label26 
         Caption         =   "INS Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   187
         Top             =   530
         Width           =   690
      End
      Begin VB.Label Label72 
         Caption         =   "Data Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   186
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label52 
         Caption         =   "Policy Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   185
         Top             =   2750
         Width           =   1020
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   120
         TabIndex        =   184
         Top             =   2430
         Width           =   975
      End
      Begin VB.Label Label54 
         Caption         =   "Case Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   183
         Top             =   3000
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   2880
         TabIndex        =   182
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Treat Reason"
         Height          =   255
         Left            =   2880
         TabIndex        =   181
         Top             =   2715
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Panel Status"
         Height          =   255
         Left            =   7440
         TabIndex        =   180
         Top             =   3990
         Width           =   1095
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   179
         Top             =   2715
         Width           =   255
      End
      Begin VB.Label Label34 
         Caption         =   "Clm Pending"
         Height          =   255
         Left            =   7440
         TabIndex        =   178
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Clinic"
         Height          =   255
         Left            =   2880
         TabIndex        =   177
         Top             =   1185
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Clinic Grp."
         Height          =   255
         Left            =   2880
         TabIndex        =   176
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   2880
         TabIndex        =   175
         ToolTipText     =   "Date Joined"
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   174
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   173
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label44 
         Caption         =   "Renewal "
         Height          =   255
         Left            =   120
         TabIndex        =   172
         Top             =   4125
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Take Over"
         Height          =   255
         Left            =   120
         TabIndex        =   171
         Top             =   4390
         Width           =   1095
      End
      Begin VB.Label Label55 
         Caption         =   "Agent Code"
         Height          =   255
         Left            =   120
         TabIndex        =   170
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label56 
         Caption         =   "Pol. Year"
         Height          =   255
         Left            =   2880
         TabIndex        =   169
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   168
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Label60 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   7440
         TabIndex        =   167
         Top             =   2680
         Width           =   855
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   166
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label62 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   7440
         TabIndex        =   165
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label38 
         Caption         =   "Payor"
         Height          =   225
         Left            =   2880
         TabIndex        =   164
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label77 
         Caption         =   "Act Amt"
         Height          =   255
         Left            =   7440
         TabIndex        =   163
         ToolTipText     =   "Approved Amt"
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label78 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   162
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label79 
         Caption         =   "App Amt"
         Height          =   255
         Left            =   7440
         TabIndex        =   161
         ToolTipText     =   "Approved Amt"
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label Label80 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   160
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label15 
         Caption         =   "MBM IC"
         Height          =   255
         Left            =   7440
         TabIndex        =   159
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   158
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   157
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "MC Days"
         Height          =   255
         Left            =   2880
         TabIndex        =   156
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   155
         Top             =   1725
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Bill No"
         Height          =   255
         Left            =   7440
         TabIndex        =   154
         Top             =   1725
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   7440
         TabIndex        =   153
         Top             =   1485
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   152
         Top             =   1485
         Width           =   255
      End
      Begin VB.Label Label33 
         Caption         =   "No of Visit"
         Height          =   255
         Left            =   2880
         TabIndex        =   151
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   150
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Label39 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Left            =   6960
         TabIndex        =   149
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   148
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   2880
         TabIndex        =   147
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label48 
         Caption         =   "Inv Date"
         Height          =   255
         Left            =   7440
         TabIndex        =   146
         Top             =   1965
         Width           =   735
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   145
         Top             =   1965
         Width           =   255
      End
      Begin VB.Label Label81 
         Caption         =   "Clinic Type"
         Height          =   255
         Left            =   7440
         TabIndex        =   143
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label69 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   142
         Top             =   3435
         Width           =   285
      End
      Begin VB.Label Label68 
         Caption         =   "OP Batch No"
         Height          =   255
         Left            =   7440
         TabIndex        =   141
         Top             =   3480
         Width           =   1005
      End
      Begin VB.Label Label63 
         Caption         =   "AL Utilized"
         Height          =   255
         Left            =   7440
         TabIndex        =   140
         ToolTipText     =   "Annual Limit Utilized"
         Top             =   3240
         Width           =   1005
      End
      Begin VB.Label Label59 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9840
         TabIndex        =   139
         Top             =   3195
         Width           =   285
      End
      Begin VB.Label Label86 
         Caption         =   "Pay Type"
         Height          =   255
         Left            =   7440
         TabIndex        =   138
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label87 
         Caption         =   "DOV Year"
         Height          =   255
         Left            =   7440
         TabIndex        =   137
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label88 
         Caption         =   "Bill Amt vs App Amt"
         Height          =   375
         Left            =   9840
         TabIndex        =   136
         Top             =   4815
         Width           =   1575
      End
      Begin VB.Label Label89 
         Caption         =   "Department"
         Height          =   255
         Left            =   2880
         TabIndex        =   135
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   2520
         TabIndex        =   201
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1215
      Left            =   120
      TabIndex        =   228
      Top             =   6840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   63
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label30 
      Caption         =   "Policy No"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   240
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label83 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   230
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label Label82 
      Caption         =   "Report Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   229
      Top             =   480
      Width           =   1065
   End
End
Attribute VB_Name = "frmClaimRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DIagCode, Clinics, Description, Payor, Day, total, PLNType, ClinicType As String
Private AA As Integer
Private strAppend As String
Private strAppendCase As String
Private AdmissionDate1, AdmissionDate2, AdmissionDate3 As String
Private AdmissionDate4, AdmissionDate5, AdmissionDate6 As String
Private AdmissionDate7, AdmissionDate8, AdmissionDate9 As String
Private AdmissionDate10, AdmissionDate11, AdmissionDate12 As String
Private AdmissionDate13, AdmissionDate14, AdmissionDate15 As String
Private AdmissionDate16, AdmissionDate17, AdmissionDate18 As String
Private AdmissionDate19, AdmissionDate20, AdmissionDate21 As String
Private AdmissionDate22, AdmissionDate23, AdmissionDate24 As String
Private AdmissionDate25, AdmissionDate26, AdmissionDate27 As String
Private AdmissionDate28, AdmissionDate29, AdmissionDate30 As String
Private AdmissionDate31 As String
Private Day1, Day2, Day3, Day4, Day5, Day6, Day7 As String
Private Day8, Day9, Day10, Day11, Day12, Day13, Day14 As String
Private Day15, Day16, Day17, Day18, Day19, Day20, Day21 As String
Private Day22, Day23, Day24, Day25, Day26, Day27, Day28 As String
Private Day29, Day30, Day31 As String
Private Total1, Total2, Total3, Total4, Total5, Total6, Total7 As String
Private Total8, Total9, Total10, Total11, Total12, Total13, Total14 As String
Private Total15, Total16, Total17, Total18, Total19, Total20, Total21 As String
Private Total22, Total23, Total24, Total25, Total26, Total27, Total28 As String
Private Total29, Total30, Total31 As String
Private intExit As Integer
Private Current As String
Private GRPCompany As String
Private Month As String
Private m_blnDirection(1 To 63) As Boolean

Private Sub cboDeptSelection_Keypress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDeptSelection.Text, cboDeptSelection.SelStart) + Chr(KeyAscii) + Mid(cboDeptSelection.Text, cboDeptSelection.SelStart + cboDeptSelection.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboDeptSelection.ListCount - 1
 
        If NewText = LCase(Left(cboDeptSelection.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDeptSelection.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboDeptSelection.Text = ValidValue
                cboDeptSelection.SelStart = 0
                cboDeptSelection.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboHLTCode_Click()
    chkPay.Value = 0
End Sub

Private Sub CboReportType_Click()
    Const METHOD_NAME = "CboReportType_Click"
    On Error GoTo debugs
    CboSelection.Clear
    
    If CboReportType.ListIndex = 0 Then 'Admission
        CboSelection.AddItem "Daily Report"
        CboSelection.AddItem "Monthly Report"
        CboSelection.AddItem "YTD Report"
        CboSelection.Text = "Daily Report"
        
        Call ListViewHeader2
    
    ElseIf CboReportType.ListIndex = 1 Then 'Ageband
        CboSelection.AddItem "Summary"
        CboSelection.AddItem "By Gender"
        CboSelection.AddItem "By Insured Type"
        CboSelection.AddItem "By Race"
        CboSelection.AddItem "By Plan Code"
        CboSelection.Text = "Summary"
        
        Call ListViewHeader3
        
    ElseIf CboReportType.ListIndex = 2 Then 'Agent No
        CboSelection.AddItem "Details Listing"
        CboSelection.AddItem "General Summary"
        CboSelection.AddItem "By Top 50 Claims"
        'CboSelection.AddItem "By Same Doctor"
        CboSelection.Text = "Details Listing"
        
        Call ListViewHeader4
        
    ElseIf CboReportType.ListIndex = 3 Then 'Bordx
        CboSelection.AddItem "Summary Bordx"
        CboSelection.AddItem "Bordx (By Grp Company)"
        CboSelection.AddItem "Bordx (By Grp Company & Department)"
        CboSelection.Text = "Summary Bordx"
        
        Call ListViewHeader19
        
    ElseIf CboReportType.ListIndex = 4 Then 'Clinic
        CboSelection.AddItem "Detail Listing"
        CboSelection.AddItem "General Summary (By Payor)"
        CboSelection.AddItem "General Summary (By Grp Company)"
        CboSelection.AddItem "General Summary (All)"
        CboSelection.AddItem "Top 20 by Claims"
        CboSelection.AddItem "Top 20 by Amount"
        CboSelection.AddItem "By Treatment Reason"
        CboSelection.AddItem "By Clinic Type"
        CboSelection.AddItem "By Clinic Type & Clinic Panel"
        CboSelection.Text = "Detail Listing"
        
        Call ListViewHeader8
        
    ElseIf CboReportType.ListIndex = 5 Then 'Clinic Group
                
        Call ListViewHeader9
    
    ElseIf CboReportType.ListIndex = 6 Then 'Department
        CboSelection.AddItem "Summary (By Visit)"
        CboSelection.AddItem "Summary (By MC)"
        CboSelection.AddItem "Employee Utilization"
        CboSelection.AddItem "Illness Utilization"
        CboSelection.AddItem "Clinic Utilization"
        CboSelection.AddItem "Department Utilization"
        CboSelection.AddItem "MC Times Analysis"
        CboSelection.AddItem "MC Days Analysis"
        CboSelection.AddItem "Visit Analysis"
        CboSelection.Text = "Summary (By Visit)"
        
        Call ListViewHeader18
    
    ElseIf CboReportType.ListIndex = 7 Then 'Diagnosis
        CboSelection.AddItem "Top 20 by Claims"
        CboSelection.AddItem "Top 20 by Amount"
        CboSelection.AddItem "General Summary"
        CboSelection.AddItem "By Clinic"
        CboSelection.Text = "General Summary"
        
        Call ListViewHeader5
    
    ElseIf CboReportType.ListIndex = 8 Then 'Gender
        CboSelection.AddItem "Summary"
        CboSelection.AddItem "By Ageband"
        CboSelection.AddItem "By Insured Type"
        CboSelection.AddItem "By Race"
        CboSelection.AddItem "By Plan Code"
        CboSelection.Text = "Summary"
        
        Call ListViewHeader7
    
    ElseIf CboReportType.ListIndex = 9 Then 'General Listing
        CboSelection.AddItem "Without Remarks"
        CboSelection.AddItem "With Remarks"
        CboSelection.AddItem "By No of Visits"
        CboSelection.Text = "Without Remarks"
        
        Call ListViewHeader1
    
    ElseIf CboReportType.ListIndex = 10 Then 'Group Company
        CboSelection.AddItem "Detail Listing"
        CboSelection.AddItem "General Summary (By Payor)"
        CboSelection.AddItem "General Summary (By Grp Company)"
        CboSelection.AddItem "Employee Utilization"
        CboSelection.AddItem "Illness Utilization"
        CboSelection.AddItem "Clinic Utilization"
        CboSelection.AddItem "Department Utilization"
        CboSelection.AddItem "MC Times Analysis"
        CboSelection.AddItem "MC Days Analysis"
        CboSelection.AddItem "Visit Analysis"
        CboSelection.AddItem "Group Company"
        CboSelection.Text = "Detail Listing"
        
        Call ListViewHeader16
    
    ElseIf CboReportType.ListIndex = 11 Then 'I/C
        
        Call ListViewHeader20
        
    ElseIf CboReportType.ListIndex = 12 Then 'Insured Type
        CboSelection.AddItem "Summary"
        CboSelection.AddItem "By Plan Code"
        CboSelection.AddItem "By Ageband"
        CboSelection.AddItem "By Gender"
        CboSelection.AddItem "By Race"
        CboSelection.Text = "Summary"
        
        Call ListViewHeader10
    
    ElseIf CboReportType.ListIndex = 13 Then 'Invoice Listing
        CboSelection.AddItem "Invoice Listing"
        CboSelection.Text = "Invoice Listing"
        
        Call ListViewHeader23
    ElseIf CboReportType.ListIndex = 14 Then 'MC
        CboSelection.AddItem "Summary"
        CboSelection.AddItem "Summary (By Clinic)"
        CboSelection.AddItem "Summary (By Grp Company)"
        CboSelection.AddItem "Summary (By Diag.)"
        CboSelection.Text = "Summary"
        
        Call ListViewHeader6
    
    ElseIf CboReportType.ListIndex = 15 Then 'Month
        CboSelection.AddItem "Eff. Month (By Payor)"
        CboSelection.AddItem "Eff. Month (All Payor)"
        CboSelection.AddItem "Exp. Month (By Payor)"
        CboSelection.AddItem "Exp. Month (All Payor)"
        CboSelection.AddItem "Visit Month (By Payor)"
        CboSelection.AddItem "Visit Month (All Payor)"
        CboSelection.AddItem "Reg. Month (By Payor)"
        CboSelection.AddItem "Reg. Month (All Payor)"
        CboSelection.Text = "Eff. Month (By Payor)"
        
        Call ListViewHeader17
        
    ElseIf CboReportType.ListIndex = 16 Then 'Monthly Reports
        CboSelection.AddItem "Utilization (By Member)"
        CboSelection.AddItem "Claims Listing"
        CboSelection.AddItem "Summary (by Clinics)"
        CboSelection.AddItem "Summary (by Illness)"
        CboSelection.AddItem "Departmental Bordx"
        CboSelection.AddItem "O/S Utilized (Combined)"
        CboSelection.AddItem "O/S Utilized (Separate)"
        CboSelection.Text = "Utilization (By Member)"
        
        Call ListViewHeader22
    
    ElseIf CboReportType.ListIndex = 17 Then 'Payor
        Call ListViewHeader11
            
    ElseIf CboReportType.ListIndex = 18 Then 'Plan Code
        CboSelection.AddItem "General Summary"
        CboSelection.AddItem "Top 10 Clinics"
        CboSelection.AddItem "Top 10 Treatment Reasons"
        CboSelection.AddItem "By Ageband"
        CboSelection.AddItem "By Insured Type"
        CboSelection.AddItem "By Insured Type & Ageband"
        CboSelection.AddItem "General Summary with Premium"
        CboSelection.AddItem "By Clinic Panel"
        CboSelection.Text = "General Summary"
        
        Call ListViewHeader12
    
    ElseIf CboReportType.ListIndex = 19 Then 'Policy
        
        Call ListViewHeader21
            
    ElseIf CboReportType.ListIndex = 20 Then 'Race
        CboSelection.AddItem "Summary"
        CboSelection.AddItem "By Ageband"
        CboSelection.AddItem "By Insured Type"
        CboSelection.AddItem "By Gender"
        CboSelection.AddItem "By Plan Code"
        CboSelection.Text = "Summary"
        
        Call ListViewHeader13
    
    ElseIf CboReportType.ListIndex = 21 Then 'State
                
        Call ListViewHeader14
    
    ElseIf CboReportType.ListIndex = 22 Then 'Utilization
        CboSelection.AddItem "Summary (All)"
        CboSelection.AddItem "Summary (By Payor)"
        CboSelection.AddItem "Summary (By Grp Company)"
        CboSelection.AddItem "Summary (By Illness)"
        CboSelection.AddItem "Summary (By Member)"
        CboSelection.AddItem "O/S Utilized"
        CboSelection.Text = "Summary (All)"
        
        Call ListViewHeader15
    
    ElseIf CboReportType.ListIndex = 23 Then 'Verification of Claims
        CboSelection.AddItem "Claims Same Date"
        CboSelection.AddItem "Claims Same Amount"
        CboSelection.Text = "Claims Same Date"
        
        Call ListViewHeader24
    
    End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboSelection_Click()
Const METHOD_NAME = "CboSelection_Click"
On Error GoTo debugs
    If CboReportType.ListIndex = 0 Then 'Admission
        Call ListViewHeader2
    ElseIf CboReportType.ListIndex = 1 Then 'Ageband
        Call ListViewHeader3
    ElseIf CboReportType.ListIndex = 2 Then 'Agent No
        Call ListViewHeader4
    ElseIf CboReportType.ListIndex = 3 Then 'Bordx
        Call ListViewHeader19
    ElseIf CboReportType.ListIndex = 4 Then 'Clinic
        Call ListViewHeader8
    ElseIf CboReportType.ListIndex = 5 Then 'Clinic Group
        Call ListViewHeader9
    ElseIf CboReportType.ListIndex = 6 Then 'Department
        Call ListViewHeader18
    ElseIf CboReportType.ListIndex = 7 Then 'Diagnosis
        Call ListViewHeader5
    ElseIf CboReportType.ListIndex = 8 Then 'Gender
        Call ListViewHeader7
    ElseIf CboReportType.ListIndex = 9 Then 'General Listing
        Call ListViewHeader1
    ElseIf CboReportType.ListIndex = 10 Then 'Group Company
        Call ListViewHeader16
    ElseIf CboReportType.ListIndex = 11 Then 'I/C
        Call ListViewHeader20
    ElseIf CboReportType.ListIndex = 12 Then 'Insured Type
        Call ListViewHeader10
    ElseIf CboReportType.ListIndex = 13 Then 'Invoice
        Call ListViewHeader23
    ElseIf CboReportType.ListIndex = 14 Then 'MC
        Call ListViewHeader6
    ElseIf CboReportType.ListIndex = 15 Then 'Month
        Call ListViewHeader17
    ElseIf CboReportType.ListIndex = 16 Then 'Monthly Reports
        Call ListViewHeader22
    ElseIf CboReportType.ListIndex = 17 Then 'Payor
        Call ListViewHeader11
    ElseIf CboReportType.ListIndex = 18 Then 'Plan Code
        Call ListViewHeader12
    ElseIf CboReportType.ListIndex = 19 Then 'Policy
        Call ListViewHeader21
    ElseIf CboReportType.ListIndex = 20 Then 'Race
        Call ListViewHeader13
    ElseIf CboReportType.ListIndex = 21 Then 'State
        Call ListViewHeader14
    ElseIf CboReportType.ListIndex = 22 Then 'Utilization
        Call ListViewHeader15
    ElseIf CboReportType.ListIndex = 23 Then 'Utilization
        Call ListViewHeader24
        
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub chkCLN_Click()
Const METHOD_NAME = "chkCLN_Click"
On Error GoTo debugs
Dim CLN As New ADODB.Recordset
Dim StrSql As String
   ' MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    cboCLNName.Clear
    If chkCLN.Value = 1 Then
        StrSql = "EXEC SearchCLN"
        CLN.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic

        Do Until CLN.EOF
            cboCLNName.AddItem CLN!CLNCode & " - " & CLN!CLNName
            CLN.MoveNext
        Loop
        CLN.Close
        Set CLN = Nothing
    End If
   ' MediConnCISWS.Close
   ' Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub chkDept2_Click()
Const METHOD_NAME = "chkDept2_Click"
On Error GoTo debugs
Dim StrSql As String
Dim DEPT As New ADODB.Recordset

  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        If chkDept2.Value = 1 Then
        
            If Len(CboGrpCom) Then
                cboDeptSelection.Clear
        
                StrSql = "Select DISTINCT MBMDepartment FROM MBMOthers where MBMGroupCompany = '" & Trim(RemStr(CboGrpCom)) & "'"
                DEPT.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            
                
                    Do While Not DEPT.EOF
                        cboDeptSelection.AddItem Trim(DEPT!MBMDepartment)
                        DEPT.MoveNext
                    Loop
                
                DEPT.Close
                Set DEPT = Nothing
            End If
        Else
            cboDeptSelection.Clear
        End If
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ChkGrpCom2_Click()
Const METHOD_NAME = "ChkGrpCom2_Click"
On Error GoTo debugs
Dim StrSql As String
Dim PAYGRP As New ADODB.Recordset
  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        CboGrpCom.Clear
        If ChkGrpCom2.Value = 1 Then
            StrSql = "select GRPCoName from GRPCompany where GRPCoPAYCode= '" & Mid(cboPayorCode, 1, 2) & "'"
            PAYGRP.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            
            
            Do While Not PAYGRP.EOF
               CboGrpCom.AddItem Trim(PAYGRP!GrpCoName)
               PAYGRP.MoveNext
            Loop
            
            PAYGRP.Close
            Set PAYGRP = Nothing
        End If
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub chkPay_Click()
Const METHOD_NAME = "chkPay_Click"
On Error GoTo debugs
Dim PAY As New ADODB.Recordset
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        cboPayorCode.Clear
        If chkPay.Value = 1 Then
            If cboPayorCode.ListCount = 0 Then
                StrSql = "Select PAYCode, PAYCompanyName from PAY Where HLTCode = '" & Mid(cboHLTCode, 1, 1) & "'"
                PAY.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                       
                Do While Not PAY.EOF
                    cboPayorCode.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
                    PAY.MoveNext
                Loop
                PAY.Close
                Set PAY = Nothing
            End If
        End If
    ''medixmasterconn.Close
    'Set medixmasterconn = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdFileLocationSUN_Click()
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
    CommonDialog1.ShowOpen
    txtImportFilenameSUN = CommonDialog1.FileTitle
    Exit Sub
  
ErrHandler:
  Exit Sub
End Sub

Private Sub cmdUploadDirect_Click()
Dim AccNo As String
Dim TransRef As String
Dim AppAmt As String
Dim AccLine As String
Dim tmpJournalNo As String
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim error As Boolean
Dim Record As ListItem
Dim lSize As Long
Dim FileExists As Integer
Dim SendDate As String
Dim tmpDebitAmt As Double
Dim tmpCreditAmt As Double
    
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim tmpPathName As String

    Screen.MousePointer = vbHourglass
        If ChkScan.Value = 0 Then
            ListView2.ListItems.Clear
            Current = 0
            total = 0
        End If
        'ChkScan.Value = 0
    If OPTClaimNo.Value <> False Or OPTBordxNo.Value <> False Then
        If txtImportFilenameSUN <> "" Then
                Val1 = 2
                error = False
                'cmpUpload.Enabled = False
                FileExists = 0
            
                lSize = -1
                lSize = FileLen(CommonDialog1.Filename)
                
                If lSize <= 0 Then
                    FileExists = 0
                ElseIf lSize > 0 Then
                    FileExists = 1
                End If
                
                txtImportFilenameSUN = CommonDialog1.Filename
                If FileExists = 1 Then
                    Set xlwbook = xl.Workbooks.Open(CommonDialog1.Filename)
                    Set xlsheet = xlwbook.Sheets.Item(1)
                    
                    Do Until Val1 = 60000
                        If xlsheet.Cells(Val1, 1) <> "" Then
                            error = False
                            TransRef = xlsheet.Cells(Val1, 1)
                            
                            Call SearchRecordExcel(TransRef)
                        Else
                            Screen.MousePointer = vbDefault
                            'cmpUpload.Enabled = True
                            Exit Sub
                        End If
                        Val1 = Val1 + 1
                    Loop
                Else
                    Screen.MousePointer = vbDefault
                    cmpUpload.Enabled = True
                    MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
                
                End If
        Else
            Screen.MousePointer = vbDefault
            cmpUpload.Enabled = True
            MsgBox "Please select file to upload!", vbOKOnly + vbCritical, DialogBoxTitle
    
        End If
    
        
     Else
            MsgBox "Please select claim number or bordx number!", vbOKOnly + vbCritical, DialogBoxTitle
            Screen.MousePointer = vbDefault
        End If

    'Screen.MousePointer = vbDefault
End Sub

Private Function SearchRecordExcel(SearchStrAppend As String)
Const METHOD_NAME = "SearchRecords"
On Error GoTo debugs
Dim sql As String
Dim strAppend As String

    strAppend = ""
    GRPCompany = ""
    Clinics = ""
    
    If OPTClaimNo.Value = True Then
        If Len(Trim(SearchStrAppend)) Then
           strAppend = strAppend + " AND CASNumber = '" & RemStr(Trim(SearchStrAppend)) & "'"
        End If
    End If
    'If Len(Trim(TxtCaseNo2)) Then
    '   strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    'End If
    If OPTBordxNo.Value = True Then
        If Len(Trim(SearchStrAppend)) Then
            strAppend = strAppend & " AND BordxRefNo = '" & Trim(SearchStrAppend) & "'"
        End If
    End If
   '
   ' If Len(Trim(TxtBordxNo2)) Then
   '     strAppend = strAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
   ' End If
    
    
    'If CboReportType.ListIndex = 0 And CboSelection.Text = "Daily Report" Then

    If CboReportType.ListIndex = 9 And CboSelection.Text = "Without Remarks" Then
        Call ResultListing(strAppend)
    'If CboReportType.ListIndex = 9 And CboSelection.Text = "With Remarks" Then
    '    Call ResultListing2(strAppend)
    'ElseIf CboReportType.ListIndex = 9 And CboSelection.Text = "By No of Visits" Then
    '    Call ResultListing3(strAppend)
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub ClaimsListingDirect(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
Dim sql2 As String
Dim rst3 As New ADODB.Recordset
Dim tmpAccName As String

    total = 0
    Current = 0
    lblTotalCR = 0
    lblTotalDR = 0
    

    
    If Trim(txtDB) = "UHC" Then
        sql = "select * from SALFLDGUHC where 0 = 0 "
    ElseIf Trim(txtDB) = "FW" Then
        sql = "select * from SALFLDGFW where 0 = 0 "
    ElseIf Trim(txtDB) = "MED" Then
        sql = "select * from SALFLDGMED where 0 = 0 "
    ElseIf Trim(txtDB) = "GV" Then
        sql = "select * from SALFLDGGV where 0 = 0 "
    ElseIf Trim(txtDB) = "TL" Then
        sql = "select * from SALFLDGTL where 0 = 0 "
    ElseIf Trim(txtDB) = "UHI" Then
        sql = "select * from SALFLDGUHI where 0 = 0 "
    ElseIf Trim(txtDB) = "MS" Then
        sql = "select * from SALFLDGMS where 0 = 0 "
    End If
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdClear.Enabled = True
        'CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        cmdExit.Enabled = True
    Else
        Do Until rst2.EOF
            With rst2
                Set Record = lstSUNListing.ListItems.Add(, , "")
                
                'If Len(!Allocation) Then
                '    Record.SubItems(1) = !Allocation
                'End If
                If Len(!ACCNT_CODE) Then
                    Record.SubItems(1) = !ACCNT_CODE
                End If
                If Len(!TRANS_DATE) Then
                    Record.SubItems(2) = Mid(!TRANS_DATE, 7, 2) & "/" & Mid(!TRANS_DATE, 5, 2) & "/" & Mid(!TRANS_DATE, 1, 4)
                End If
                If Len(!TREFERENCE) Then
                    Record.SubItems(3) = Trim(!TREFERENCE)
                End If
                If Len(!Amount) Then
                    'If !Amount < 0 Then
                     '   Record.SubItems(4) = !Amount * -1
                    'Else
                        Record.SubItems(4) = Format(!Amount, "#,##0.00")
                    'End If
                    If !D_C = "C" Then
                        lblTotalCR = lblTotalCR + !Amount
                        lblTotalCR = Format(lblTotalCR, "#,##0.00")
                    Else
                        lblTotalDR = lblTotalDR + !Amount
                        lblTotalDR = Format(lblTotalDR, "#,##0.00")
                    End If
                End If
                If Len(!D_C) Then
                    Record.SubItems(5) = !D_C
                End If
                If Len(!Allocation) Then
                    Record.SubItems(6) = !Allocation
                End If
                If Len(!Period) Then
                   Record.SubItems(7) = !Period
                End If
                If Len(!DESCRIPTN) Then
                    Record.SubItems(8) = !DESCRIPTN
                End If
                If Len(!JRNAL_TYPE) Then
                    Record.SubItems(9) = !JRNAL_TYPE
                End If
                
                If Len(!JRNAL_NO) Then
                    Record.SubItems(10) = !JRNAL_NO
                End If
        
                If Len(!JRNAL_SRCE) Then
                   Record.SubItems(11) = !JRNAL_SRCE
                End If
                
                If Len(!JRNAL_LINE) Then
                   Record.SubItems(12) = !JRNAL_LINE
                End If
                If Len(!ALLOC_REF) Then
                    Record.SubItems(13) = !ALLOC_REF
                End If
                If Len(!ALLOC_DATE) Then
                    If Trim(!ALLOC_DATE) <> "0" Then
                        Record.SubItems(14) = Mid(!ALLOC_DATE, 7, 2) & "/" & Mid(!ALLOC_DATE, 5, 2) & "/" & Mid(!ALLOC_DATE, 1, 4)
                    End If
                End If
                If Len(!ANAL_T1) Then
                    Record.SubItems(15) = !ANAL_T1
                End If
                If Len(!ANAL_T2) Then
                    Record.SubItems(16) = !ANAL_T2
                End If
                If Len(!ANAL_T3) Then
                    Record.SubItems(17) = !ANAL_T3
                End If
                If Len(!ANAL_T4) Then
                    Record.SubItems(18) = !ANAL_T4
                End If
                If Len(!ANAL_T5) Then
                    Record.SubItems(19) = !ANAL_T5
                End If
                
                If Len(!ANAL_T6) Then
                    Record.SubItems(20) = !ANAL_T6
                End If
                
                If Len(!ANAL_T7) Then
                    Record.SubItems(21) = !ANAL_T7
                End If
                If Len(!ANAL_T8) Then
                    Record.SubItems(22) = !ANAL_T8
                End If
            
                If Len(!ANAL_T9) Then
                    Record.SubItems(23) = !ANAL_T9
                End If
            
                If Trim(txtDB) = "UHC" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'UHC' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "FW" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'FW' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "MED" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'MED' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "GV" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'GV' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "TL" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'TL' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "UHI" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'UHI' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                ElseIf Trim(txtDB) = "MS" Then
                    sql2 = " SELECT SUN_DB, ACCNT_CODE,ACCNT_NAME  from SSRFACC WHERE SUN_DB = 'MS' And ACCNT_CODE = '" & Trim(!ACCNT_CODE) & "'"
                End If
                    
                rst3.Open sql2, medixmasterSUNSys, adOpenKeyset, adLockOptimistic
                    If rst3.EOF = False Then
                        tmpAccName = IIf(Len(rst3!ACCNT_NAME), rst3!ACCNT_NAME, "")
                    End If
                rst3.Close
                Set rst3 = Nothing
                
                If Len(tmpAccName) Then
                    Record.SubItems(24) = Trim(tmpAccName)
                End If
            
            End With
            rst2.MoveNext
        Loop
    End If

    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    cmdExit.Enabled = True
    
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

Private Sub cmdBack_Click()
Const METHOD_NAME = "cmdBack_Click"
On Error GoTo debugs
    ListView2.Height = 855
    ListView2.Left = 0
    ListView2.Top = 6800
    ListView2.Width = 11775
    CboReportType.Visible = True
    CboSelection.Visible = True
    Frame22.Visible = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdStop_Click()
Const METHOD_NAME = "CmdStop_Click"
On Error GoTo debugs
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdView_Click()
Const METHOD_NAME = "CmdView_Click"
On Error GoTo debugs
    ListView2.Height = 6615
    ListView2.Left = 0
    ListView2.Top = 480
    ListView2.Width = 11775
    CboReportType.Visible = False
    CboSelection.Visible = False
    Frame22.Visible = False
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ListViewHeader1()
Const METHOD_NAME = "ListViewHeader1"
On Error GoTo debugs
Label47.Caption = "DOV"

If CboSelection.Text = "Without Remarks" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Claim No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "BordxNo"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Bordx Date"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Member No"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Policy No"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Covered ID"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Prin. Name"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "Patient Name"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(10).Text = "Prin. IC No"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(11).Text = "Prin. Sex"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Prin. Ageband"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Prin. DOB"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Patient IC No"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = "Patient Sex"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Text = "Patient Ageband"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Patient DOB"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(18).Text = "Insured Type"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Text = "Plan Code"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Text = "Clinic Code"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(21).Text = "Clinic Name"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(22).Text = "Treatment Code"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Text = "Description"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(24).Text = "Bill Amt"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "Approved Amt"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Recovery Amt"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Premium Amt"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Case Status"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Text = "Claimability"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Text = "Policy Status"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Text = "Data Status"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Text = "DOV"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Text = "DOR"
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Text = "MC Days"
    ListView2.ColumnHeaders(34).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(35).Text = "Worksheet Date"
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Text = "Payor Eff. Date"
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(37).Text = "Payor Exp. Date"
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Text = "Department"
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Text = "PayType"
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Text = "GP Amt"
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Text = "SP Amt"
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Text = "Insured Amt"
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ''MURALI
    ListView2.ColumnHeaders(43).Text = "Invoice No"
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(44).Text = "Group Company"
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(45).Text = "Diagnosis Desc"
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(46).Text = "Diagnosis Desc2"
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(47).Text = "Diagnosis Desc3"
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(48).Text = "Diagnosis Desc4"
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(49).Text = "Diagnosis Desc5"
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(50).Text = "Register By"
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(51).Text = "LastUpdate By"
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    
    For AA = 52 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "With Remarks" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Claim No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "BordxNo"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Bordx Date"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Member No"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Policy No"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Covered ID"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Prin. Name"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "Patient Name"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(10).Text = "Prin. IC No"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(11).Text = "Prin. Sex"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Prin. Ageband"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Prin. DOB"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Patient IC No"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = "Patient Sex"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Text = "Patient Ageband"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Patient DOB"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(18).Text = "Insured Type"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Text = "Plan Code"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Text = "Clinic Code"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(21).Text = "Clinic Name"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(22).Text = "Treatment Code"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Text = "Description"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(24).Text = "Bill Amt"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "Approved Amt"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Recovery Amt"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Premium Amt"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Case Status"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Text = "Claimability"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Text = "Policy Status"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Text = "Data Status"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Text = "DOV"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Text = "DOR"
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Text = "MC Days"
    ListView2.ColumnHeaders(34).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(35).Text = "Worksheet Date"
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Text = "Payor Eff. Date"
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(37).Text = "Payor Exp. Date"
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Text = "Remarks"
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    
    For AA = 39 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By No of Visits" Then
    ListView2.ColumnHeaders(1).Text = "No of Visits"
    ListView2.ColumnHeaders(2).Text = "Patient IC No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Patient Name"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Plan Code"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Amt Claims"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ListViewHeader2()
Const METHOD_NAME = "ListViewHeader2"
On Error GoTo debugs
    If CboSelection.Text = "Daily Report" Then
        ListView2.ColumnHeaders(1).Text = "No"
        ListView2.ColumnHeaders(2).Text = "Claim No"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "Member No"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Policy No"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Patient Name"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "DOV"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(7).Text = "Clinic Code"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(8).Text = "Clinic Name"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(9).Text = "Treatment Code"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(10).Text = "MC Days"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
        
        For AA = 11 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf CboSelection.Text = "Monthly Report" Then
        ListView2.ColumnHeaders(1).Text = "Payor"
        ListView2.ColumnHeaders(2).Text = "Date 1"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "Total"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(4).Text = "Date 2"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Total"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(6).Text = "Date 3"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Total"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(8).Text = "Date 4"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(9).Text = "Total"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(10).Text = "Date 5"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(11).Text = "Total"
        ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(12).Text = "Date 6"
        ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(13).Text = "Total"
        ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(14).Text = "Date 7"
        ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(15).Text = "Total"
        ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(16).Text = "Date 8"
        ListView2.ColumnHeaders(16).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(17).Text = "Total"
        ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(18).Text = "Date 9"
        ListView2.ColumnHeaders(18).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(19).Text = "Total"
        ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(20).Text = "Date 10"
        ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(21).Text = "Total"
        ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(22).Text = "Date 11"
        ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(23).Text = "Total"
        ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(24).Text = "Date 12"
        ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(25).Text = "Total"
        ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(26).Text = "Date 13"
        ListView2.ColumnHeaders(26).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(27).Text = "Total"
        ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(28).Text = "Date 14"
        ListView2.ColumnHeaders(28).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(29).Text = "Total"
        ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(30).Text = "Date 15"
        ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(31).Text = "Total"
        ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(32).Text = "Date 16"
        ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(33).Text = "Total"
        ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(34).Text = "Date 17"
        ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(35).Text = "Total"
        ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(36).Text = "Date 18"
        ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(37).Text = "Total"
        ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(38).Text = "Date 19"
        ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(39).Text = "Total"
        ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(40).Text = "Date 20"
        ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(41).Text = "Total"
        ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(42).Text = "Date 21"
        ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(43).Text = "Total"
        ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(44).Text = "Date 22"
        ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(45).Text = "Total"
        ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(46).Text = "Date 23"
        ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(47).Text = "Total"
        ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(48).Text = "Date 24"
        ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(49).Text = "Total"
        ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(50).Text = "Date 25"
        ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(51).Text = "Total"
        ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(52).Text = "Date 26"
        ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(53).Text = "Total"
        ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(54).Text = "Date 27"
        ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(55).Text = "Total"
        ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(56).Text = "Date 28"
        ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(57).Text = "Total"
        ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(58).Text = "Date 29"
        ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(59).Text = "Total"
        ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(60).Text = "Date 30"
        ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(61).Text = "Total"
        ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(62).Text = "Date 31"
        ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(63).Text = "Total"
        ListView2.ColumnHeaders(63).Alignment = lvwColumnRight
    
    ElseIf CboSelection.Text = "YTD Report" Then
        ListView2.ColumnHeaders(1).Text = "Payor"
        ListView2.ColumnHeaders(2).Text = "January"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(3).Text = "February"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(4).Text = "March"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(5).Text = "April"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(6).Text = "May"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(7).Text = "June"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(8).Text = "July"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(9).Text = "August"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(10).Text = "September"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(11).Text = "October"
        ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(12).Text = "November"
        ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(13).Text = "December"
        ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
        
        For AA = 14 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
        
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ListViewHeader3()
   Const METHOD_NAME = "ListViewHeader3"
On Error GoTo debugs
 
If CboSelection.Text = "Summary" Then
    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Gender" Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Insured Type" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Race" Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Plan Code" Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Age Band"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
'ElseIf Option30.Value = True Then
    'ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    'ListView2.ColumnHeaders(2).Text = "Age Band"
    'ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    'ListView2.ColumnHeaders(3).Text = "No of Visit"
    'ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    'ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    'ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    'For AA = 6 To 63
        'ListView2.ColumnHeaders(AA).Text = ""
    'Next AA

End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ListViewHeader4()
 Const METHOD_NAME = "ListViewHeader4"
On Error GoTo debugs
If CboSelection.Text = "Details Listing" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Claim No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Member No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Covered ID"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Doctor"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Bill Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Approved Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Recovery Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "Treatment Code"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Description"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "DOV"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "MC Days"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "Ins Code"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Text = "Plan Code"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Case Status"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    For AA = 18 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "General Summary" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "By Top 50 Claims" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA


ElseIf CboSelection.Text = "By Same Doctor" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Doctor Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ListViewHeader9()

    ListView2.ColumnHeaders(1).Text = "Clinic Group"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Bill Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub ListViewHeader8()

If CboSelection.Text = "General Summary (By Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Top 20 by Amount" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Top 20 by Claims" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "General Summary (All)" Then
    ListView2.ColumnHeaders(1).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Text = "Clinic Name"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Location"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "No of Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "MC Times"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "MC Days"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Bill Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Approved Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Average Bill"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "Average Approved"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "%MC"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "Days Per MC"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    For AA = 14 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "By Treatment Reason" Then
    ListView2.ColumnHeaders(1).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Text = "Clinic Name"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Treatment Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Description"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "No of Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Clinic Type" Then
    ListView2.ColumnHeaders(1).Text = "Clinic Type"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Amt Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    
    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Detail Listing" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Claim No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Member No"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Covered ID"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Patient Name"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Bill Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Approved Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Member Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "DOV"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Case Status"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Claimability"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Treatment Code"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Text = "MC Days"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    
    For AA = 16 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "By Clinic Type & Clinic Panel" Then
    ListView2.ColumnHeaders(1).Text = "Clinic Type"
    ListView2.ColumnHeaders(2).Text = "No of Claims (Panel)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Amt Claims (Panel)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "No of Claims (Non-Panel)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Amt Claims (Non-Panel)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total No of Claims"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Amt Claims"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "General Summary (By Grp Company)" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End If

End Sub

Private Sub ListViewHeader7()
      
If CboSelection.Text = "Summary" Then
    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Ageband" Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Insured Type" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Race" Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Gender"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Plan Code" Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
'ElseIf Option30.Value = True Then
    'ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    'ListView2.ColumnHeaders(2).Text = "Gender"
    'ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    'ListView2.ColumnHeaders(3).Text = "No of Visit"
    'ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    'ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    'ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    'For AA = 6 To 63
        'ListView2.ColumnHeaders(AA).Text = ""
    'Next AA
End If

End Sub

Private Sub ListViewHeader13()

If CboSelection.Text = "Summary" Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
        
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "By Ageband" Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Insured Type" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Gender" Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Plan Code" Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End If
End Sub

Private Sub ListViewHeader10()

If CboSelection.Text = "Summary" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Ins Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "By Plan Code" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "By Ageband" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Gender" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Race" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If

End Sub

Private Sub ListViewHeader12()
    
If CboSelection.Text = "General Summary" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "General Summary with Premium" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Basic Premium Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Premium Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Top 10 Clinics" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Clinic Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Clinic Name"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "No of Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Top 10 Treatment Reasons" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Treatment Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Descriptions"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "No of Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Ageband" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Insured Type" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "By Insured Type & Ageband" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Ageband"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "No of Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "By Clinic Panel" Then
    ListView2.ColumnHeaders(1).Text = "Plan No"
    ListView2.ColumnHeaders(2).Text = "Claim Amt"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Panel Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Non-Panel Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
        
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    

End If

End Sub

Private Sub ListViewHeader14()

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "State"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End Sub

Private Sub ListViewHeader15()
        
If CboSelection.Text = "Summary (By Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Grp Company Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Patient Name"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Plan Code"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "No of Visit"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "MC Days"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Actual Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Approved Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    
    For AA = 10 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Summary (All)" Then
    ListView2.ColumnHeaders(1).Text = "Member No"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "MC Days"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Summary (By Grp Company)" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "O/S Utilized" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Covered ID"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Member Name"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Total Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "MC Times"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "MC Days"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Claim Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Summary (By Illness)" Then
    ListView2.ColumnHeaders(1).Text = "Treatment Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visits (GP)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "No of Visits (SP)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "No of Visits (Others)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Visits"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "GP Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "SP Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Others Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Non GP Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "Total Bill Amt"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "Average Cost (GP)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "Average Cost (SP)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "Average Cost (Others)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "Total Average Cost"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    
    For AA = 17 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Summary (By Member)" Then
    ListView2.ColumnHeaders(1).Text = "Member No"
    ListView2.ColumnHeaders(2).Text = "Pricipal Name"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total Visits"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Times"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "MC Days"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Annual Limit"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Balance Limit"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
        
    For AA = 10 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA


End If

End Sub

Private Sub ListViewHeader16()

If CboSelection.Text = "General Summary (By Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Group Company"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "General Summary (By Grp Company)" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Bill Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Average Bill"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Approved"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Detail Listing" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Group Company"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Claim No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Member No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Covered ID"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Member Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "DOV"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "Case Status"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Claimability"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Treatment Code"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "MC Days"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    
    For AA = 15 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Employee Utilization" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Department"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Mem_No."
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Name"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Staff No."
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Ann_Lmt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "App_Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Bal Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "O/S Utilization (%)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    
    For AA = 10 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Illness Utilization" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Department"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Illness"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Times"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "MC Days"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "App Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Clinic Utilization" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Department"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Times"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "MC Days"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "App Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Department Utilization" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Department"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Times"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Days"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "App Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "MC Times Analysis" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Department"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "MC Times"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Days"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "App Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "MC Days Analysis" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Department"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "MC Days"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Times"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "App Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Visit Analysis" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Department"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Times"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Days"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "App Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Group Company" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Department"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Times"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Days"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "App Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End If

End Sub

Private Sub ListViewHeader17()

If CboSelection.Text = "Eff. Month (By Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Eff. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Eff. Month (All Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Eff. Month"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "MC Days"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Exp. Month (By Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Exp. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Exp. Month (All Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Exp. Month"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "MC Days"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Visit Month (By Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Visit Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Visit Month (All Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Visit Month"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "MC Days"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Reg. Month (By Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Reg. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Reg. Month (All Payor)" Then
    ListView2.ColumnHeaders(1).Text = "Reg. Month"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "MC Days"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End If

End Sub

Private Sub ListViewHeader18()

If CboSelection.Text = "Summary (By Visit)" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Summary (By MC)" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "MC Days"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "No of MC"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Employee Utilization" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Mem_No."
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Staff No."
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Ann_Lmt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "App_Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Bal Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "O/S Utilization (%)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Illness Utilization" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Illness"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Times"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Days"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "App Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Clinic Utilization" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Clinic"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Times"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "MC Days"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "App Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Department Utilization" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "MC Times"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "App Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "MC Times Analysis" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "MC Times"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "App Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "MC Days Analysis" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "MC Days"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Times"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "App Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Visit Analysis" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "MC Times"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "App Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End If

End Sub

Private Sub ListViewHeader19()

If CboSelection.Text = "Summary Bordx" Then
    ListView2.ColumnHeaders(1).Text = "Bordx No"
    ListView2.ColumnHeaders(2).Text = "Bordx Date"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Bordx Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
       ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Bordx (By Grp Company)" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "Bordx No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Bordx Date"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Bordx Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
       ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Bordx (By Grp Company & Department)" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Group Company"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Bordx No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Bordx Date"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "No of Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Bordx Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        
    For AA = 9 To 63
       ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
    
End Sub

Private Sub ListViewHeader20()

    ListView2.ColumnHeaders(1).Text = "I/C No"
    ListView2.ColumnHeaders(2).Text = "Patient Name"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "O/S Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
            
    For AA = 7 To 63
       ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End Sub

Private Sub ListViewHeader21()

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Policy No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Group Company"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
       ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub ListViewHeader22()

If CboSelection.Text = "Utilization (By Member)" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Cover ID"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Member Name"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "MC Times"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "MC Days"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Claim Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "GP Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "SP Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "Insurance Claims"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(12).Text = "FFS"
    'ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    
    For AA = 12 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Claims Listing" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Claim No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Bordx No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bordx Date"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Member No"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Cover ID"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Patient Name"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Clinic Code"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Clinic Name"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "PayType"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "Treatment Code 1"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Description 1"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "Bill Amt"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "GP Amt"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "SP Amt"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "Insured Amt"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "DOV"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "MC Days"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    
    ListView2.ColumnHeaders(19).Text = "Treatment Code 2"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "Description 2"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "Treatment Code 3"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "Description 3"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    
    ListView2.ColumnHeaders(23).Text = "Treatment Code 4"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "Description 4"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "Treatment Code 5"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Description 5"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Summary (by Clinics)" Then
    ListView2.ColumnHeaders(1).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Text = "Clinic Name"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "No of Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "MC Times"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "MC Days"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Bill"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "% MCs"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Days per MC"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    
    
    For AA = 11 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Summary (by Illness)" Then
    ListView2.ColumnHeaders(1).Text = "Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "GP Visits"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "SP Visits"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Visits"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "GP Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "SP Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Insurance Claims"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Bill Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "%"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "Avg GP Cost"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Avg SP Cost"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "Avg Cost"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    
    For AA = 14 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Departmental Bordx" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Bordx No."
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Bordx Date"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "GP Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "SP Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Insurance Claims"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "O/S Utilized (Combined)" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member Name (Principal)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Total Visit"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "MC Times"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "MC Days"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Annual limit"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Topup Amount"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Claim Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "FFS"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "Balance Limit"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Percentage Utilized (%)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    
    For AA = 13 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "O/S Utilized (Separate)" Then
    ListView2.ColumnHeaders(1).Text = "Department"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Cover ID"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Patient Name"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Total Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "MC Times"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "MC Days"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Annual limit"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Topup Amount"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Claim Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "FFS"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Balance Limit"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "Percentage Utilized (%)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    
    For AA = 13 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If

End Sub

Private Sub ListViewHeader6()

If CboSelection.Text = "Summary" Then
    ListView2.ColumnHeaders(1).Text = "MC Days"
    ListView2.ColumnHeaders(2).Text = "No of MC"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Actual Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Summary (By Clinic)" Then
    ListView2.ColumnHeaders(1).Text = "Clinic Code"
    ListView2.ColumnHeaders(2).Text = "Clinic Name"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "MC Days"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "No of MC"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Summary (By Grp Company)" Then
    ListView2.ColumnHeaders(1).Text = "Group Company"
    ListView2.ColumnHeaders(2).Text = "MC Days"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "No of MC"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Summary (By Diag.)" Then
    ListView2.ColumnHeaders(1).Text = "Treatment Code"
    ListView2.ColumnHeaders(2).Text = "Descriptions"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "MC Days"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "No of MC"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End If

End Sub


Private Sub ListViewHeader11()

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "No of Visit"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End Sub

Private Sub Form_Load()
    Const METHOD_NAME = "Form_Load"
    On Error GoTo debugs
    intExit = 0
    Call ComboListing
    
    CboReportType.ListIndex = 9
    CboSelection.Text = "Without Remarks"
        
    Call ListViewHeader1
    
    ListView2.Visible = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

' =================== Command Button Click =============================

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdPrint_Click()
Const METHOD_NAME = "CmdPrint_Click"
On Error GoTo debugs
    Screen.MousePointer = vbHourglass
    
    If ListView2.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(ListView2)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
   Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs
Dim strAppend As String
    strAppend = ""
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    cmdExit.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
        'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
            Call SearchRecords
        'MediConnCISWS.Close
        'Set MediConnCISWS = Nothing
    
    Screen.MousePointer = vbDefault
    
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Function ClearForm(frmCall)
On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    ElseIf TypeOf ctl Is checkbox Then
            ctl.Value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    
    End If
Next ctl
End Function

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = ""
    ListView2.ListItems.Clear
End Sub
' =================== Command Button Click / Change ==========================


'============== ComboBox Listing =================================
Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim PolStatus As New ADODB.Recordset
Dim state As New ADODB.Recordset
Dim ClinicGrp As New ADODB.Recordset
    
    'PolStatus.Open "Select pstatuscode ,pstatusDesc from PolicyStatus", medixmasterconn, adOpenForwardOnly, adLockReadOnly
    'state.Open "Select STADescription from STA", medixmasterconn, adOpenForwardOnly, adLockReadOnly
    'ClinicGrp.Open "Select ClinicGRPCode, ClinicGRPName from ClinicGRP", medixmasterconn, adOpenForwardOnly, adLockReadOnly
    'MediConnCISWS.Open "File Name=" & BOSPath & "Conn\MediConnCISWS.UDL"
    
    cboHLTCode.AddItem "S" & " - " & "Sihat"
    cboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    
    CboInsType.AddItem "I" & " - " & "Individual"
    CboInsType.AddItem "F" & " - " & "Family"
    CboInsType.AddItem "G" & " - " & "Group Individual"
    CboInsType.AddItem "H" & " - " & "Group Family"
    
    CboIEinv.AddItem "Yes"
    
    CboMemType.AddItem "Both"
    CboMemType.AddItem "P" & " - " & "Principal"
    CboMemType.AddItem "S" & " - " & "Supplementary"
    CboMemType.Text = "Both"
    
    CboRace.AddItem "C - Chinese"
    CboRace.AddItem "I - Indian"
    CboRace.AddItem "M - Malay"
    CboRace.AddItem "O - Other"
            
    CboSex.AddItem "M" & " - " & "Male"
    CboSex.AddItem "F" & " - " & "Female"
    
    CboRelation.AddItem "D - Daughter"
    CboRelation.AddItem "H - Husband"
    CboRelation.AddItem "S - Son"
    CboRelation.AddItem "W - Wife"
    
    cboDataStatus.AddItem "A - Active"
    cboDataStatus.AddItem "D - Delete"
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SearchPolStatus"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    PolStatus.CursorLocation = adUseClient
    PolStatus.CursorType = adOpenForwardOnly
    PolStatus.CacheSize = 100
    Set PolStatus = cmd.Execute
    
    Do Until PolStatus.EOF
        CboStatus.AddItem (PolStatus!pstatuscode) & " - " & (PolStatus!pstatusDesc)
        PolStatus.MoveNext
    Loop
    PolStatus.Close
    Set PolStatus = Nothing
        
    With CboCaseStatus
        .AddItem "0 - Open"
        .AddItem "1 - Open & Close"
        .AddItem "2 - Delete"
        .AddItem "3 - All"
        .Text = "1 - Open & Close"
    End With
    
    CboExpired.AddItem "N - No"
    CboExpired.AddItem "Y - Yes"
    
    cboPanelStatus.AddItem "C - Confirmed"
    cboPanelStatus.AddItem "Y - Yes"
    
    CboClmPending.AddItem "N - No"
    CboClmPending.AddItem "Y - Yes"
        
    CboPreMode.AddItem "A - Annually"
    CboPreMode.AddItem "Q - Quarterly"
    CboPreMode.AddItem "S - Semi Annual"
    CboPreMode.AddItem "M - Monthly"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SearchState"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    state.CursorLocation = adUseClient
    state.CursorType = adOpenForwardOnly
    state.CacheSize = 100
    Set state = cmd.Execute
    
    Do Until state.EOF
        CboState.AddItem ChkStr(state!STADescription)
        state.MoveNext
    Loop
    state.Close
    Set state = Nothing
        
    cboRenewalInd.AddItem "N - New"
    cboRenewalInd.AddItem "P - Partial"
    cboRenewalInd.AddItem "R - Renew"
    cboRenewalInd.AddItem "T - Timely"
    
    cboTakeOver.AddItem "N - None"
    cboTakeOver.AddItem "Y - Take Over"
    cboTakeOver.AddItem "C - Conversion"
              
    cboSuppType.AddItem "Both"
    cboSuppType.AddItem "C" & " - " & "Combined Limit"
    cboSuppType.AddItem "S" & " - " & "Separated Limit"
    cboSuppType.Text = "Both"
        
    CboClinicType.AddItem "0 - Specialised"
    CboClinicType.AddItem "1 - General Practisional"
    CboClinicType.AddItem "2 - Others"
    
    CboPayType.AddItem "C - Clinic"
    CboPayType.AddItem "M - Member"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SearchCLNGrp"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    ClinicGrp.CursorLocation = adUseClient
    ClinicGrp.CursorType = adOpenForwardOnly
    ClinicGrp.CacheSize = 100
    Set ClinicGrp = cmd.Execute
    
    Do Until ClinicGrp.EOF
        cboCLNGrp.AddItem Trim(ClinicGrp!ClinicGrpName)
        ClinicGrp.MoveNext
    Loop
    ClinicGrp.Close
    Set ClinicGrp = Nothing
    
    'MediConnCISWS.Close
    'Set MediConnCISWS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'============== ComboBox Listing =================================


'==================== List Final Diag. Listing Resullts ==========================
Private Sub FinalDiagListing(Optional strAppend As String)
Const METHOD_NAME = "FinalDiagListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim StrSql As String

    DIagCode = ""
    lblCurrent = 0
    ListView2.ListItems.Clear
            
    If CboSelection.Text = "Top 20 by Claims" Then
        strAppendCase = "104"
        'sql = " SELECT Top 20 COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved, " & _
              " Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays " & _
              " From VIEW_CLMReport Where 0 = 0 " & strAppend & _
              " GROUP BY CASTreatReason1 ORDER BY Total DESC"
    
    ElseIf CboSelection.Text = "Top 20 by Amount" Then
        strAppendCase = "105"
        'sql = "SELECT Top 20 Sum(CLMTotalApproved) As TotalApproved, COUNT(CASTreatReason1) AS Total, CASTreatReason1, " & _
              " Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays " & _
              " From VIEW_CLMReport Where 0 = 0 " & strAppend & _
              " Group By CASTreatReason1 Order By TotalApproved Desc "
              
    ElseIf CboSelection.Text = "General Summary" Then
        strAppendCase = "106"
        'sql = "SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved, " & _
              " Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays " & _
              " From VIEW_CLMReport Where 0 = 0 " & strAppend & _
              " GROUP BY CASTreatReason1 ORDER BY Total DESC"
    End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
       
        Do Until rst2.EOF
    
            With rst2
                
                If Len(rst2!CASTreatReason1) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CASTreatReason1)
                    DIagCode = rst2!CASTreatReason1
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    DIagCode = ""
                End If
                
                'sql2 = " SELECT ReasonCode, ReasonDesc From TreatmentReason Where ReasonCode = '" & Trim(DIagCode) & "'"
                'rst3.Open sql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
                StrSql = "EXEC CISClaimTracking '" & Trim(DIagCode) & "',10"
                rst3.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                       
                If rst3.EOF = True Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(1) = ""
                Else
                    Description = Trim(rst3!ReasonDesc)
                    Record.SubItems(1) = Trim(Description)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(rst2!total) Then
                    Record.SubItems(2) = rst2!total
                End If
                
                If Len(rst2!MCDays) Then
                    Record.SubItems(3) = rst2!MCDays
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(4) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
        
    If ListView2.ListItems.Count = 0 Then
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

Private Sub FinalDiagbyHos(Optional strAppend As String)
Const METHOD_NAME = "FinalDiagbyHos"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim StrSql As String
    
    DIagCode = ""
    lblCurrent = 0
    
    ListView2.ListItems.Clear
       
    'sql = " SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual, ClinicCode, CLNName, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 " & strAppend & _
          " GROUP BY CASTreatReason1, ClinicCode, CLNName "
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "103"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
       
        Do Until rst2.EOF
    
            With rst2
                               
                If Len(rst2!CASTreatReason1) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CASTreatReason1)
                    DIagCode = rst2!CASTreatReason1
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    DIagCode = ""
                End If
                
                'sql2 = " SELECT ReasonCode, ReasonDesc From TreatmentReason Where ReasonCode = '" & Trim(DIagCode) & "'"
                'rst3.Open sql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly 'adOpenKeyset, adLockOptimistic
                StrSql = "EXEC CISClaimTracking '" & Trim(DIagCode) & "',10"
                rst3.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                       
                If rst3.EOF = True Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(1) = ""
                Else
                    Description = Trim(rst3!ReasonDesc)
                    Record.SubItems(1) = Trim(Description)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(rst2!ClinicCode) Then
                    Record.SubItems(2) = Trim(rst2!ClinicCode)
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(rst2!CLNName) Then
                    Record.SubItems(3) = Trim(rst2!CLNName)
                Else
                    Record.SubItems(3) = "Empty"
                End If
                
                If Len(rst2!total) Then
                    Record.SubItems(4) = rst2!total
                End If
                
                If Len(rst2!MCDays) Then
                    Record.SubItems(5) = rst2!MCDays
                End If
                                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(6) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(7) = Format(rst2!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
        
    If ListView2.ListItems.Count = 0 Then
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

'================ Daily Visit Listing ===========================
Private Sub DailyVisitListing(Optional strAppend As String)
Const METHOD_NAME = "DailyVisitListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim intloop As Integer
       
    lblCurrent = 0
    ListView2.ListItems.Clear
            
    'sql = " SELECT CASNumber, MBMNumber, MBMPolicyNo, " & _
          " PatientName, DateVisit, ClinicCode, CLNName, CASTreatReason1, MCDays " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "34"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                    
            With rst2
            
                intloop = intloop + 1
                
                Set Record = ListView2.ListItems.Add(, , intloop)
                                                
                If Len(rst2!CasNumber) Then
                       Record.SubItems(1) = Trim(rst2!CasNumber)
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPolicyNo) Then
                    Record.SubItems(3) = rst2!MBMPolicyNo
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!DateVisit) Then
                    Record.SubItems(5) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                
                If Len(rst2!ClinicCode) Then
                    Record.SubItems(6) = Trim(rst2!ClinicCode)
                End If
                
                If Len(rst2!CLNName) Then
                    Record.SubItems(7) = Trim(rst2!CLNName)
                End If
                
                If Len(rst2!CASTreatReason1) Then
                    Record.SubItems(8) = Trim(rst2!CASTreatReason1)
                End If
                
                If Len(rst2!MCDays) Then
                    Record.SubItems(9) = Trim(rst2!MCDays)
                End If
            End With
            rst2.MoveNext
            lblCurrent = lblCurrent + 1
        Loop
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'================ Daily Visit Listing ===========================

'=========== Option Button, Combo Box Change/KeyPress ===================

Private Sub cboPayorCode_Change()

    If cboPayorCode.Text = "" Then
        If InStr(cboPayorCode.Text, "null") Then
           cboPayorCode.Enabled = False
        End If
    End If
End Sub

Private Sub cboPayorCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayorCode.Text, cboPayorCode.SelStart) + Chr(KeyAscii) + Mid(cboPayorCode.Text, cboPayorCode.SelStart + cboPayorCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayorCode.ListCount - 1
 
        If NewText = LCase(Left(cboPayorCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayorCode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayorCode.Text = ValidValue
                cboPayorCode.SelStart = 0
                cboPayorCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode.Text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.Text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHLTCode.ListCount - 1
 
        If NewText = LCase(Left(cboHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTCode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHLTCode.Text = ValidValue
                cboHLTCode.SelStart = 0
                cboHLTCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboInsType_Change()

    If CboInsType.Text = "" Then
        If InStr(CboInsType.Text, "null") Then
           CboInsType.Enabled = False
        End If
    End If
End Sub

Private Sub CboInsType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInsType.Text, CboInsType.SelStart) + Chr(KeyAscii) + Mid(CboInsType.Text, CboInsType.SelStart + CboInsType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboInsType.ListCount - 1
 
        If NewText = LCase(Left(CboInsType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboInsType.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboInsType.Text = ValidValue
                CboInsType.SelStart = 0
                CboInsType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboMemType_Change()

    If CboMemType.Text = "" Then
        If InStr(CboMemType.Text, "null") Then
           CboMemType.Enabled = False
        End If
    End If
End Sub

Private Sub cboSuppType_Change()

    If cboSuppType.Text = "" Then
        If InStr(cboSuppType.Text, "null") Then
           cboSuppType.Enabled = False
        End If
    End If
End Sub

Private Sub CboRace_Change()

    If CboRace.Text = "" Then
        If InStr(CboRace.Text, "null") Then
           CboRace.Enabled = False
        End If
    End If
End Sub

Private Sub CboRace_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRace.Text, CboRace.SelStart) + Chr(KeyAscii) + Mid(CboRace.Text, CboRace.SelStart + CboRace.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboRace.ListCount - 1
 
        If NewText = LCase(Left(CboRace.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRace.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboRace.Text = ValidValue
                CboRace.SelStart = 0
                CboRace.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboSex_Change()

    If CboSex.Text = "" Then
        If InStr(CboSex.Text, "null") Then
           CboSex.Enabled = False
        End If
    End If
End Sub

Private Sub CboSex_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSex.Text, CboSex.SelStart) + Chr(KeyAscii) + Mid(CboSex.Text, CboSex.SelStart + CboSex.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboSex.ListCount - 1
 
        If NewText = LCase(Left(CboSex.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSex.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboSex.Text = ValidValue
                CboSex.SelStart = 0
                CboSex.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboRelation_Change()

    If CboRelation.Text = "" Then
        If InStr(CboRelation.Text, "null") Then
           CboRelation.Enabled = False
        End If
    End If
End Sub

Private Sub CboRelation_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRelation.Text, CboRelation.SelStart) + Chr(KeyAscii) + Mid(CboRelation.Text, CboRelation.SelStart + CboRelation.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboRelation.ListCount - 1
 
        If NewText = LCase(Left(CboRelation.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRelation.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboRelation.Text = ValidValue
                CboRelation.SelStart = 0
                CboRelation.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub cboDataStatus_Change()

    If cboDataStatus.Text = "" Then
        If InStr(cboDataStatus.Text, "null") Then
           cboDataStatus.Enabled = False
        End If
    End If
End Sub

Private Sub cboDataStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDataStatus.Text, cboDataStatus.SelStart) + Chr(KeyAscii) + Mid(cboDataStatus.Text, cboDataStatus.SelStart + cboDataStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboDataStatus.ListCount - 1
 
        If NewText = LCase(Left(cboDataStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDataStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboDataStatus.Text = ValidValue
                cboDataStatus.SelStart = 0
                cboDataStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClaimability_Change()

    If CboClaimability.Text = "" Then
        If InStr(CboClaimability.Text, "null") Then
           CboClaimability.Enabled = False
        End If
    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClaimability.Text, CboClaimability.SelStart) + Chr(KeyAscii) + Mid(CboClaimability.Text, CboClaimability.SelStart + CboClaimability.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClaimability.ListCount - 1
 
        If NewText = LCase(Left(CboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClaimability.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClaimability.Text = ValidValue
                CboClaimability.SelStart = 0
                CboClaimability.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboStatus_Change()

    If CboStatus.Text = "" Then
        If InStr(CboStatus.Text, "null") Then
           CboStatus.Enabled = False
        End If
    End If
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboStatus.Text, CboStatus.SelStart) + Chr(KeyAscii) + Mid(CboStatus.Text, CboStatus.SelStart + CboStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboStatus.ListCount - 1
 
        If NewText = LCase(Left(CboStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboStatus.Text = ValidValue
                CboStatus.SelStart = 0
                CboStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCaseStatus_Change()

    If CboCaseStatus.Text = "" Then
        If InStr(CboCaseStatus.Text, "null") Then
           CboCaseStatus.Enabled = False
        End If
    End If
End Sub

Private Sub CboCaseStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCaseStatus.Text, CboCaseStatus.SelStart) + Chr(KeyAscii) + Mid(CboCaseStatus.Text, CboCaseStatus.SelStart + CboCaseStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCaseStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCaseStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCaseStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCaseStatus.Text = ValidValue
                CboCaseStatus.SelStart = 0
                CboCaseStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub CboExpired_Change()

    If CboExpired.Text = "" Then
        If InStr(CboExpired.Text, "null") Then
           CboExpired.Enabled = False
        End If
    End If
End Sub

Private Sub CboExpired_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboExpired.Text, CboExpired.SelStart) + Chr(KeyAscii) + Mid(CboExpired.Text, CboExpired.SelStart + CboExpired.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboExpired.ListCount - 1
 
        If NewText = LCase(Left(CboExpired.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboExpired.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboExpired.Text = ValidValue
                CboExpired.SelStart = 0
                CboExpired.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPreMode_Change()

    If CboPreMode.Text = "" Then
        If InStr(CboPreMode.Text, "null") Then
           CboPreMode.Enabled = False
        End If
    End If
End Sub

Private Sub CboPreMode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPreMode.Text, CboPreMode.SelStart) + Chr(KeyAscii) + Mid(CboPreMode.Text, CboPreMode.SelStart + CboPreMode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPreMode.ListCount - 1
 
        If NewText = LCase(Left(CboPreMode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPreMode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPreMode.Text = ValidValue
                CboPreMode.SelStart = 0
                CboPreMode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboState_Change()

    If CboState.Text = "" Then
        If InStr(CboState.Text, "null") Then
           CboState.Enabled = False
        End If
    End If
End Sub

Private Sub CboState_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboState.Text, CboState.SelStart) + Chr(KeyAscii) + Mid(CboState.Text, CboState.SelStart + CboState.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboState.ListCount - 1
 
        If NewText = LCase(Left(CboState.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboState.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboState.Text = ValidValue
                CboState.SelStart = 0
                CboState.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboRenewalInd_Change()

    If cboRenewalInd.Text = "" Then
        If InStr(cboRenewalInd.Text, "null") Then
           cboRenewalInd.Enabled = False
        End If
    End If
End Sub

Private Sub cboRenewalInd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRenewalInd.Text, cboRenewalInd.SelStart) + Chr(KeyAscii) + Mid(cboRenewalInd.Text, cboRenewalInd.SelStart + cboRenewalInd.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboRenewalInd.ListCount - 1
 
        If NewText = LCase(Left(cboRenewalInd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRenewalInd.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboRenewalInd.Text = ValidValue
                cboRenewalInd.SelStart = 0
                cboRenewalInd.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboTakeOver_Change()

    If cboTakeOver.Text = "" Then
        If InStr(cboTakeOver.Text, "null") Then
           cboTakeOver.Enabled = False
        End If
    End If
End Sub

Private Sub cboTakeOver_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboTakeOver.Text, cboTakeOver.SelStart) + Chr(KeyAscii) + Mid(cboTakeOver.Text, cboTakeOver.SelStart + cboTakeOver.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboTakeOver.ListCount - 1
 
        If NewText = LCase(Left(cboTakeOver.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboTakeOver.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboTakeOver.Text = ValidValue
                cboTakeOver.SelStart = 0
                cboTakeOver.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboGrpCom_Change()

    If CboGrpCom.Text = "" Then
        If InStr(CboGrpCom.Text, "null") Then
           CboGrpCom.Enabled = False
        End If
    End If
End Sub

Private Sub CboGrpCom_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpCom.Text, CboGrpCom.SelStart) + Chr(KeyAscii) + Mid(CboGrpCom.Text, CboGrpCom.SelStart + CboGrpCom.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboGrpCom.ListCount - 1
 
        If NewText = LCase(Left(CboGrpCom.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpCom.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboGrpCom.Text = ValidValue
                CboGrpCom.SelStart = 0
                CboGrpCom.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNName_Change()

    If cboCLNName.Text = "" Then
        If InStr(cboCLNName.Text, "null") Then
           cboCLNName.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNName_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNName.Text, cboCLNName.SelStart) + Chr(KeyAscii) + Mid(cboCLNName.Text, cboCLNName.SelStart + cboCLNName.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNName.ListCount - 1
 
        If NewText = LCase(Left(cboCLNName.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNName.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNName.Text = ValidValue
                cboCLNName.SelStart = 0
                cboCLNName.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNGrp_Change()

    If cboCLNGrp.Text = "" Then
        If InStr(cboCLNGrp.Text, "null") Then
           cboCLNGrp.Enabled = False
        End If
    End If
End Sub

Private Sub cboCLNGrp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNGrp.Text, cboCLNGrp.SelStart) + Chr(KeyAscii) + Mid(cboCLNGrp.Text, cboCLNGrp.SelStart + cboCLNGrp.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNGrp.ListCount - 1
 
        If NewText = LCase(Left(cboCLNGrp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNGrp.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNGrp.Text = ValidValue
                cboCLNGrp.SelStart = 0
                cboCLNGrp.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClinicType_Change()

    If CboClinicType.Text = "" Then
        If InStr(CboClinicType.Text, "null") Then
           CboClinicType.Enabled = False
        End If
    End If
End Sub

Private Sub CboClinicType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClinicType.Text, CboClinicType.SelStart) + Chr(KeyAscii) + Mid(CboClinicType.Text, CboClinicType.SelStart + CboClinicType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClinicType.ListCount - 1
 
        If NewText = LCase(Left(CboClinicType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClinicType.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClinicType.Text = ValidValue
                CboClinicType.SelStart = 0
                CboClinicType.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)

If CboReportType.ListIndex = 0 And CboSelection.Text = "Daily Report" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        End Select
        
    ElseIf CboReportType.ListIndex = 0 And CboSelection.Text = "Monthly Report" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(30)
        End Select
        
    ElseIf CboReportType.ListIndex = 0 And CboSelection.Text = "YTD Report" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        End Select
        
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "Summary" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
        
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "By Gender" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
        
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "By Insured Type" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "By Race" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
    
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "By Plan Code" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 2 And CboSelection.Text = "Details Listing" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(17)
        End Select
    
    ElseIf CboReportType.ListIndex = 2 And CboSelection.Text = "General Summary" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
    
    ElseIf CboReportType.ListIndex = 2 And CboSelection.Text = "By Top 50 Claims" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
            
    ElseIf CboReportType.ListIndex = 3 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "Top 20 by Amount" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
    
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "Top 20 by Claims" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        End Select
    
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "General Summary (By Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
    
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "General Summary (All)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        End Select
    
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "Detail Listing" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        End Select
    
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "By Treatment Reason" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
    
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "By Clinic Type" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        End Select
    
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "By Clinic Type && Clinic Panel" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
    
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "General Summary (By Grp Company)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
        
    ElseIf CboReportType.ListIndex = 5 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
        
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Summary (By MC)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
    
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Utilisation" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
        
    ElseIf CboReportType.ListIndex = 7 And CboSelection.Text = "Top 20 by Claims" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 7 And CboSelection.Text = "Top 20 by Amount" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 7 And CboSelection.Text = "General Summary" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 7 And CboSelection.Text = "By Clinic" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
    
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "Summary" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
    
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "By Ageband" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
    
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "By Insured Type" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "By Race" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
    
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "By Plan Code" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 9 And CboSelection.Text = "Without Remarks" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(37)
        End Select
    
    ElseIf CboReportType.ListIndex = 9 And CboSelection.Text = "With Remarks" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(37)
        Case 38
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(38)
        End Select
    
    ElseIf CboReportType.ListIndex = 9 And CboSelection.Text = "By No of Visits" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "General Summary (By Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
    
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "General Summary (By Grp Company)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "Detail Listing" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        End Select
        
    ElseIf CboReportType.ListIndex = 11 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "Summary" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
    
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "By Plan Code" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "By Ageband" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "By Gender" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "By Race" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf (CboReportType.ListIndex = 13 And CboSelection.Text = "Summary") Or _
        (CboReportType.ListIndex = 16 And CboSelection.Text = "By Clinic Panel") Or _
        (CboReportType.ListIndex = 6 And CboSelection.Text = "Summary (By Visit)") Then
        
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
    
    ElseIf CboReportType.ListIndex = 13 And CboSelection.Text = "Invoice Listing" Then
            Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
            
        End Select
        
    ElseIf (CboReportType.ListIndex = 13 And CboSelection.Text = "Summary (By Clinic)") Or _
        (CboReportType.ListIndex = 13 And CboSelection.Text = "Summary (By Grp Company)") Or _
        (CboReportType.ListIndex = 20 And CboSelection.Text = "Summary (By Grp Company)") Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
        
    ElseIf CboReportType.ListIndex = 13 And CboSelection.Text = "Summary (By Diag.)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Eff. Month (By Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
    
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Eff. Month (All Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
    
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Exp. Month (By Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
    
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Exp. Month (All Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
    
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Visit Month (By Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
    
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Visit Month (All Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
    
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Reg. Month (By Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
    
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Reg. Month (All Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
    
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Utilization (By Member)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        End Select
    
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Claims Listing" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
        End Select
        
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Summary (by Clinics)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        End Select
        
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Summary (by Illness)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        End Select
        
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Departmental Bordx" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
    
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "O/S Utilized (Combined)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        End Select
    
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "O/S Utilized (Separate)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        End Select
        
    ElseIf CboReportType.ListIndex = 16 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
        
    ElseIf CboReportType.ListIndex = 17 And CboSelection.Text = "General Summary with Premium" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
        
    ElseIf CboReportType.ListIndex = 17 And CboSelection.Text = "General Summary" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
        
    ElseIf CboReportType.ListIndex = 17 And CboSelection.Text = "Top 10 Clinics" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
        
    ElseIf CboReportType.ListIndex = 17 And CboSelection.Text = "Top 10 Treatment Reasons" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
        
    ElseIf CboReportType.ListIndex = 17 And CboSelection.Text = "By Ageband" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    
    ElseIf CboReportType.ListIndex = 17 And CboSelection.Text = "By Insured Type" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 17 And CboSelection.Text = "By Insured Type & Ageband" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
            
    ElseIf CboReportType.ListIndex = 18 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 19 And CboSelection.Text = "Summary" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
    
    ElseIf CboReportType.ListIndex = 19 And CboSelection.Text = "By Ageband" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
        
    ElseIf CboReportType.ListIndex = 19 And CboSelection.Text = "By Insured Type" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 19 And CboSelection.Text = "By Gender" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
        
    ElseIf CboReportType.ListIndex = 19 And CboSelection.Text = "By Plan Code" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 20 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
        
    ElseIf CboReportType.ListIndex = 21 And CboSelection.Text = "Summary (By Payor)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        End Select
    
    ElseIf CboReportType.ListIndex = 21 And CboSelection.Text = "Summary (All)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
    
    ElseIf CboReportType.ListIndex = 21 And CboSelection.Text = "Summary (By Grp Company)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf CboReportType.ListIndex = 21 And CboSelection.Text = "O/S Utilized" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        End Select
        
    ElseIf CboReportType.ListIndex = 21 And CboSelection.Text = "Summary (By Illness)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        End Select
        
    ElseIf CboReportType.ListIndex = 21 And CboSelection.Text = "Summary (By Member)" Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        End Select
        
    End If

End Sub
'=========== Combo Box Change/KeyPress ===================

'=========== Check Date Format ===========================

Private Sub TxtDOV1_LostFocus()
    If IsDate(TxtDOV1) Then
    Else
        If TxtDOV1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOV1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOV2_LostFocus()
    If IsDate(TxtDOV2) Then
    Else
        If TxtDOV2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOV2.SetFocus
        End If
    End If
End Sub


Private Sub TxtDOR1_LostFocus()
    If IsDate(TxtDOR1) Then
    Else
        If TxtDOR1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOR1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOR2_LostFocus()
    If IsDate(TxtDOR2) Then
    Else
        If TxtDOR2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOR2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOW1_LostFocus()
    If IsDate(txtDOW1) Then
    Else
        If txtDOW1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOW1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOW2_LostFocus()
    If IsDate(txtDOW2) Then
    Else
        If txtDOW2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOW2.SetFocus
        End If
    End If
End Sub

Private Sub TxtMCDay1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtMCDay2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtOSUtilized1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtOSUtilized2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtPayEffDate1_LostFocus()
    If IsDate(TxtPayEffDate1) Then
    Else
        If TxtPayEffDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayEffDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayEffDate2_LostFocus()
    If IsDate(TxtPayEffDate2) Then
    Else
        If TxtPayEffDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayEffDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayExpDate1_LostFocus()
    If IsDate(TxtPayExpDate1) Then
    Else
        If TxtPayExpDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayExpDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayExpDate2_LostFocus()
    If IsDate(TxtPayExpDate2) Then
    Else
        If TxtPayExpDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayExpDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateJoin1_LostFocus()
    If IsDate(txtDateJoin1) Then
    Else
        If txtDateJoin1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateJoin1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateJoin2_LostFocus()
    If IsDate(txtDateJoin2) Then
    Else
        If txtDateJoin2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateJoin2.SetFocus
        End If
    End If
End Sub

Private Sub txtDOB1_LostFocus()
    If IsDate(txtDOB1) Then
    Else
        If txtDOB1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB1.SetFocus
        End If
    End If
End Sub

Private Sub txtDOB2_LostFocus()
    If IsDate(txtDOB2) Then
    Else
        If txtDOB2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB2.SetFocus
        End If
    End If
End Sub
'=========== Check Date Format ===========================
Private Sub ListViewHeader23()
        
    If CboSelection.Text = "Invoice Listing" Then
        ListView2.ColumnHeaders(1).Text = "Clinic Code"
        ListView2.ColumnHeaders(2).Text = "Clinic Name"
        ListView2.ColumnHeaders(3).Text = "Cases"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(4).Text = "Invoice No"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Invoice Date"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "Process month"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Invoice Amount"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(8).Text = "PIC"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(9).Text = "OP Batch No"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(10).Text = "GP/SP/NP"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft

        For AA = 11 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
    End If
End Sub
'================ Monthy Reports Listing ===========================
Private Sub MonthyReportListing(Optional strAppend As String)
Const METHOD_NAME = "MonthyReportListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
       
    ListView2.ListItems.Clear
    lblCurrent = 0
            
    'sql = " SELECT Distinct PAYCode From VIEW_CLMReport Where 0 = 0 "
    'sql2 = " Order by PAYCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "35"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                    
            With rst2
                If Len(rst2!PAYCode) Then
                
                Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Payor = rst2!PAYCode
                
                Call SearchVisitRecords(strAppend)
                                
                If Len(AdmissionDate1) And Len(Day1) Then
                    Record.SubItems(1) = AdmissionDate1 & "-" & Day1
                End If
                
                If Len(Total1) Then
                    Record.SubItems(2) = Total1
                End If
                
                If Len(AdmissionDate2) And Len(Day2) Then
                    Record.SubItems(3) = AdmissionDate2 & "-" & Day2
                End If
                
                If Len(Total2) Then
                    Record.SubItems(4) = Total2
                End If
                
                If Len(AdmissionDate3) And Len(Day3) Then
                    Record.SubItems(5) = AdmissionDate3 & "-" & Day3
                End If
                
                If Len(Total3) Then
                    Record.SubItems(6) = Total3
                End If
                
                If Len(AdmissionDate4) And Len(Day4) Then
                    Record.SubItems(7) = AdmissionDate4 & "-" & Day4
                End If
                
                If Len(Total4) Then
                    Record.SubItems(8) = Total4
                End If
                
                If Len(AdmissionDate5) And Len(Day5) Then
                    Record.SubItems(9) = AdmissionDate5 & "-" & Day5
                End If
                
                If Len(Total5) Then
                    Record.SubItems(10) = Total5
                End If
                
                If Len(AdmissionDate6) And Len(Day6) Then
                    Record.SubItems(11) = AdmissionDate6 & "-" & Day6
                End If
                
                If Len(Total6) Then
                    Record.SubItems(12) = Total6
                End If
                
                If Len(AdmissionDate7) And Len(Day7) Then
                    Record.SubItems(13) = AdmissionDate7 & "-" & Day7
                End If
                
                If Len(Total7) Then
                    Record.SubItems(14) = Total7
                End If
                
                If Len(AdmissionDate8) And Len(Day8) Then
                    Record.SubItems(15) = AdmissionDate8 & "-" & Day8
                End If
                
                If Len(Total8) Then
                    Record.SubItems(16) = Total8
                End If
                
                If Len(AdmissionDate9) And Len(Day9) Then
                    Record.SubItems(17) = AdmissionDate9 & "-" & Day9
                End If
                
                If Len(Total9) Then
                    Record.SubItems(18) = Total9
                End If
                
                If Len(AdmissionDate10) And Len(Day10) Then
                    Record.SubItems(19) = AdmissionDate10 & "-" & Day10
                End If
                
                If Len(Total10) Then
                    Record.SubItems(20) = Total10
                End If
                
                If Len(AdmissionDate11) And Len(Day11) Then
                    Record.SubItems(21) = AdmissionDate11 & "-" & Day11
                End If
                
                If Len(Total11) Then
                    Record.SubItems(22) = Total11
                End If
                
                If Len(AdmissionDate12) And Len(Day12) Then
                    Record.SubItems(23) = AdmissionDate12 & "-" & Day12
                End If
                
                If Len(Total12) Then
                    Record.SubItems(24) = Total12
                End If
                
                If Len(AdmissionDate13) And Len(Day13) Then
                    Record.SubItems(25) = AdmissionDate13 & "-" & Day13
                End If
                
                If Len(Total13) Then
                    Record.SubItems(26) = Total13
                End If
                
                If Len(AdmissionDate14) And Len(Day14) Then
                    Record.SubItems(27) = AdmissionDate14 & "-" & Day14
                End If
                
                If Len(Total14) Then
                    Record.SubItems(28) = Total14
                End If
                
                If Len(AdmissionDate15) And Len(Day15) Then
                    Record.SubItems(29) = AdmissionDate15 & "-" & Day15
                End If
                
                If Len(Total15) Then
                    Record.SubItems(30) = Total15
                End If
                
                If Len(AdmissionDate16) And Len(Day16) Then
                    Record.SubItems(31) = AdmissionDate16 & "-" & Day16
                End If
                
                If Len(Total16) Then
                    Record.SubItems(32) = Total6
                End If
                
                If Len(AdmissionDate17) And Len(Day17) Then
                    Record.SubItems(33) = AdmissionDate17 & "-" & Day17
                End If
                
                If Len(Total17) Then
                    Record.SubItems(34) = Total17
                End If
                
                If Len(AdmissionDate18) And Len(Day18) Then
                    Record.SubItems(35) = AdmissionDate18 & "-" & Day18
                End If
                
                If Len(Total18) Then
                    Record.SubItems(36) = Total8
                End If
                
                If Len(AdmissionDate19) And Len(Day19) Then
                    Record.SubItems(37) = AdmissionDate19 & "-" & Day19
                End If
                
                If Len(Total19) Then
                    Record.SubItems(38) = Total19
                End If
                
                If Len(AdmissionDate20) And Len(Day20) Then
                    Record.SubItems(39) = AdmissionDate20 & "-" & Day20
                End If
                
                If Len(Total20) Then
                    Record.SubItems(40) = Total20
                End If
                
                If Len(AdmissionDate21) And Len(Day21) Then
                    Record.SubItems(41) = AdmissionDate21 & "-" & Day21
                End If
                
                If Len(Total21) Then
                    Record.SubItems(42) = Total21
                End If
                
                If Len(AdmissionDate22) And Len(Day22) Then
                    Record.SubItems(43) = AdmissionDate22 & "-" & Day22
                End If
                
                If Len(Total22) Then
                    Record.SubItems(44) = Total22
                End If
                
                If Len(AdmissionDate23) And Len(Day23) Then
                    Record.SubItems(45) = AdmissionDate23 & "-" & Day23
                End If
                
                If Len(Total23) Then
                    Record.SubItems(46) = Total23
                End If
                
                If Len(AdmissionDate24) And Len(Day24) Then
                    Record.SubItems(47) = AdmissionDate24 & "-" & Day24
                End If
                
                If Len(Total24) Then
                    Record.SubItems(48) = Total24
                End If
                
                If Len(AdmissionDate25) And Len(Day24) Then
                    Record.SubItems(49) = AdmissionDate25 & "-" & Day25
                End If
                
                If Len(Total25) Then
                    Record.SubItems(50) = Total25
                End If
                
                If Len(AdmissionDate26) And Len(Day26) Then
                    Record.SubItems(51) = AdmissionDate26 & "-" & Day26
                End If
                
                If Len(Total26) Then
                    Record.SubItems(52) = Total26
                End If
                
                If Len(AdmissionDate27) And Len(Day27) Then
                    Record.SubItems(53) = AdmissionDate27 & "-" & Day27
                End If
                
                If Len(Total27) Then
                    Record.SubItems(54) = Total27
                End If
                
                If Len(AdmissionDate28) And Len(Day28) Then
                    Record.SubItems(55) = AdmissionDate28 & "-" & Day28
                End If
                
                If Len(Total28) Then
                    Record.SubItems(56) = Total28
                End If
                
                If Len(AdmissionDate29) And Len(Day29) Then
                    Record.SubItems(57) = AdmissionDate29 & "-" & Day29
                End If
                
                If Len(Total29) Then
                    Record.SubItems(58) = Total29
                End If
                
                If Len(AdmissionDate30) And Len(Day30) Then
                    Record.SubItems(59) = AdmissionDate30 & "-" & Day30
                End If
                
                If Len(Total30) Then
                    Record.SubItems(60) = Total30
                End If
                
                If Len(AdmissionDate31) And Len(Day31) Then
                    Record.SubItems(61) = AdmissionDate31 & "-" & Day31
                End If
                
                If Len(Total31) Then
                    Record.SubItems(62) = Total31
                End If
                End If
            End With
            rst2.MoveNext
            lblCurrent = lblCurrent + 1
        Loop
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
'================ Admission Listing ===========================

'==================== Search One Month Visit Records ===========================
Private Sub SearchVisitRecords(Optional strAppend As String)
Const METHOD_NAME = "SearchVisitRecords"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim strAppend1 As String

    ListView3.ListItems.Clear
    Call ClearBoolean
       
    strAppend1 = strAppend + " AND PAYCode = '" & Trim(Payor) & "'"
    'sql = " Select PAYCode, DateVisit, Count(CASNumber) AS Total From VIEW_CLMReport Where PAYCode = '" & Trim(Payor) & "'"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + (" Group By PAYCode, DateVisit Order by DateVisit")
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "96"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend1)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
                
        Do Until rst2.EOF
        
            With rst2
            
                TxtDay.Text = ""
                txtDate.Text = ""
                
                Set Record = ListView3.ListItems.Add(, , rst2!PAYCode)
                
                If one = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate1 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day1 = Day
                        Record.SubItems(1) = AdmissionDate1 & "-" & Day1
                        Record.SubItems(2) = rst2!total
                        Total1 = rst2!total
                        one = True
                    End If
                
                ElseIf two = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate2 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day2 = Day
                        Record.SubItems(1) = AdmissionDate2 & "-" & Day2
                        Record.SubItems(2) = rst2!total
                        Total2 = rst2!total
                        two = True
                    End If
                
                ElseIf three = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate3 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day3 = Day
                        Record.SubItems(1) = AdmissionDate3 & "-" & Day3
                        Record.SubItems(2) = rst2!total
                        Total3 = rst2!total
                        three = True
                    End If
                
                ElseIf four = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate4 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day4 = Day
                        Record.SubItems(1) = AdmissionDate4 & "-" & Day4
                        Record.SubItems(2) = rst2!total
                        Total4 = rst2!total
                        four = True
                    End If
                ElseIf five = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate5 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day5 = Day
                        Record.SubItems(1) = AdmissionDate5 & "-" & Day5
                        Record.SubItems(2) = rst2!total
                        Total5 = rst2!total
                        five = True
                    End If
                ElseIf six = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate6 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day6 = Day
                        Record.SubItems(1) = AdmissionDate6 & "-" & Day6
                        Record.SubItems(2) = rst2!total
                        Total6 = rst2!total
                        six = True
                    End If
                ElseIf seven = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate7 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day7 = Day
                        Record.SubItems(1) = AdmissionDate7 & "-" & Day7
                        Record.SubItems(2) = rst2!total
                        Total7 = rst2!total
                        seven = True
                    End If
                ElseIf eight = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate8 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day8 = Day
                        Record.SubItems(1) = AdmissionDate8 & "-" & Day8
                        Record.SubItems(2) = rst2!total
                        Total8 = rst2!total
                        eight = True
                    End If
                ElseIf nine = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate9 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day9 = Day
                        Record.SubItems(1) = AdmissionDate9 & "-" & Day9
                        Record.SubItems(2) = rst2!total
                        Total9 = rst2!total
                        nine = True
                    End If
                ElseIf ten = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate10 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day10 = Day
                        Record.SubItems(1) = AdmissionDate10 & "-" & Day10
                        Record.SubItems(2) = rst2!total
                        Total10 = rst2!total
                        ten = True
                    End If
                ElseIf eleven = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate11 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day11 = Day
                        Record.SubItems(1) = AdmissionDate11 & "-" & Day11
                        Record.SubItems(2) = rst2!total
                        Total11 = rst2!total
                        eleven = True
                    End If
                ElseIf twelve = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate12 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day12 = Day
                        Record.SubItems(1) = AdmissionDate12 & "-" & Day12
                        Record.SubItems(2) = rst2!total
                        Total12 = rst2!total
                        twelve = True
                    End If
                ElseIf thirteen = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate13 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day13 = Day
                        Record.SubItems(1) = AdmissionDate13 & "-" & Day13
                        Record.SubItems(2) = rst2!total
                        Total13 = rst2!total
                        thirteen = True
                    End If
                ElseIf fourteen = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate14 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day14 = Day
                        Record.SubItems(1) = AdmissionDate14 & "-" & Day14
                        Record.SubItems(2) = rst2!total
                        Total14 = rst2!total
                        fourteen = True
                    End If
                ElseIf fiftteen = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate15 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day15 = Day
                        Record.SubItems(1) = AdmissionDate15 & "-" & Day15
                        Record.SubItems(2) = rst2!total
                        Total15 = rst2!total
                        fiftteen = True
                    End If
                ElseIf sixteen = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate16 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day16 = Day
                        Record.SubItems(1) = AdmissionDate16 & "-" & Day16
                        Record.SubItems(2) = rst2!total
                        Total16 = rst2!total
                        sixteen = True
                    End If
                ElseIf seventeen = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate17 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day17 = Day
                        Record.SubItems(1) = AdmissionDate17 & "-" & Day17
                        Record.SubItems(2) = rst2!total
                        Total17 = rst2!total
                        seventeen = True
                    End If
                ElseIf eightteen = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate18 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day18 = Day
                        Record.SubItems(1) = AdmissionDate18 & "-" & Day18
                        Record.SubItems(2) = rst2!total
                        Total18 = rst2!total
                        eightteen = True
                    End If
                ElseIf nineteen = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate19 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day19 = Day
                        Record.SubItems(1) = AdmissionDate19 & "-" & Day19
                        Record.SubItems(2) = rst2!total
                        Total19 = rst2!total
                        nineteen = True
                    End If
                ElseIf twenty = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate20 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate20 & "-" & Day20
                        Record.SubItems(2) = rst2!total
                        Total20 = rst2!total
                        twenty = True
                    End If
                ElseIf twentyone = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate21 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day21 = Day
                        Record.SubItems(1) = AdmissionDate21 & "-" & Day21
                        Record.SubItems(2) = rst2!total
                        Total21 = rst2!total
                        twentyone = True
                    End If
                ElseIf twentytwo = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate22 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day22 = Day
                        Record.SubItems(1) = AdmissionDate22 & "-" & Day22
                        Record.SubItems(2) = rst2!total
                        Total22 = rst2!total
                        twentytwo = True
                    End If
                ElseIf twentythree = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate23 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate23 & "-" & Day23
                        Record.SubItems(2) = rst2!total
                        Total23 = rst2!total
                        twentythree = True
                    End If
                ElseIf twentyfour = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate24 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day24 = Day
                        Record.SubItems(1) = AdmissionDate24 & "-" & Day24
                        Record.SubItems(2) = rst2!total
                        Total24 = rst2!total
                        twentyfour = True
                    End If
                ElseIf twentyfive = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate25 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day25 = Day
                        Record.SubItems(1) = AdmissionDate25 & "-" & Day25
                        Record.SubItems(2) = rst2!total
                        Total25 = rst2!total
                        twentyfive = True
                    End If
                ElseIf twentysix = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate26 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day26 = Day
                        Record.SubItems(1) = AdmissionDate26 & "-" & Day26
                        Record.SubItems(2) = rst2!total
                        Total26 = rst2!total
                        twentysix = True
                    End If
                ElseIf twentyseven = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate27 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day27 = Day
                        Record.SubItems(1) = AdmissionDate27 & "-" & Day27
                        Record.SubItems(2) = rst2!total
                        Total27 = rst2!total
                        twentyseven = True
                    End If
                ElseIf twentyeight = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate28 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day28 = Day
                        Record.SubItems(1) = AdmissionDate28 & "-" & Day28
                        Record.SubItems(2) = rst2!total
                        Total28 = rst2!total
                        twentyeight = True
                    End If
                ElseIf twentynine = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate29 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day29 = Day
                        Record.SubItems(1) = AdmissionDate29 & "-" & Day29
                        Record.SubItems(2) = rst2!total
                        Total29 = rst2!total
                        twentynine = True
                    End If
                ElseIf thirty = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate30 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day30 = Day
                        Record.SubItems(1) = AdmissionDate30 & "-" & Day30
                        Record.SubItems(2) = rst2!total
                        Total30 = rst2!total
                        thirty = True
                    End If
                ElseIf thirtyone = False Then
                    If Len(rst2!DateVisit) Then
                        Record.SubItems(1) = rst2!DateVisit
                        AdmissionDate31 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day31 = Day
                        Record.SubItems(1) = AdmissionDate31 & "-" & Day31
                        Record.SubItems(2) = rst2!total
                        Total31 = rst2!total
                        thirtyone = True
                    End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'==================== Search One Month Visit Records ===========================

'==================== Check Day ===========================
Private Sub CheckDay()
Const METHOD_NAME = "CheckDay"
On Error GoTo debugs
Dim intYears As Integer 'Used as a counter
Dim intNumDaysToCheck As Long 'Used to limit search
Dim tmpDate As Date 'Temporary date variable
Dim tmpMonth As Date 'Temporary month variable
Dim intPaydays As Integer 'Counts number of paydays Each month
Dim intPayDayOfWeek As Integer 'Pay Day of Week
Dim intX As Long 'Counter
    'computes number of days between pay date and 10 years from paydate
    intNumDaysToCheck = DateDiff("d", CDate(txtDate), DateAdd("yyyy", 100, CDate(txtDate)))
    intPayDayOfWeek = Weekday(CDate(txtDate))
    intPaydays = 0
    tmpMonth = DatePart("m", CDate(txtDate))
    'get paid every two weeks
    'if you get paid every friday and want to check for
    '5 paycheck months, change this to step 7 and change
    'commented value below


    For intX = 0 To intNumDaysToCheck Step 14
        'set the day we are working with
        tmpDate = CDate(txtDate) + intX
        'make sure the month hasn't changed


        If tmpMonth = DatePart("m", tmpDate) Then
            'if it is a payday


            If Weekday(tmpDate) <> intPayDayOfWeek Then
                'there is a problem
                MsgBox "An Error was encountered. Calculation cannot continue"
            Else
                'increment the paydays
                intPaydays = intPaydays + 1
                'if you have arrived at a 3 payday month, add it to the list
                'if you get weekly paychecks, change this value from 3 to 5
            End If
        Else
            'change the month
            tmpMonth = DatePart("m", tmpDate)
            'set the number of paydays to 1
            intPaydays = 1
        End If
    Next
    
    Call DateDay
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'==================== Check Day ===========================

' ================= Call Day ==============================
Private Sub DateDay()
Const METHOD_NAME = "DateDay"
On Error GoTo debugs
    Select Case Weekday(CDate(txtDate))
        Case vbMonday
        TxtDay = "Monday"
        Case vbTuesday
        TxtDay = "Tuesday"
        Case vbWednesday
        TxtDay = "Wednesday"
        Case vbThursday
        TxtDay = "Thursday"
        Case vbFriday
        TxtDay = "Friday"
        Case vbSaturday
        TxtDay = "Saturday"
        Case vbSunday
        TxtDay = "Sunday"
        Case Else
        TxtDay = "Error With date"
    End Select
    
    Day = TxtDay
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
' ================= Call Day ==============================

'=================== Clear String Records ========================
Public Function ClearBoolean()
    one = False
    two = False
    three = False
    four = False
    five = False
    six = False
    seven = False
    eight = False
    nine = False
    ten = False
    eleven = False
    twelve = False
    thirteen = False
    fourteen = False
    fifteen = False
    sixteen = False
    seventeen = False
    eightteen = False
    nineteen = False
    twenty = False
    twentyone = False
    twentytwo = False
    twentythree = False
    twentyfour = False
    twentyfive = False
    twentysix = False
    twentyseven = False
    twentyeight = False
    twentynine = False
    thirty = False
    thirtyone = False
    
    AdmissionDate1 = ""
    Day1 = ""
    Total1 = ""
    
    AdmissionDate2 = ""
    Day2 = ""
    Total2 = ""
    
        
    AdmissionDate3 = ""
    Day3 = ""
    Total3 = ""
    
    AdmissionDate4 = ""
    Day4 = ""
    Total4 = ""
    
    AdmissionDate5 = ""
    Day5 = ""
    Total5 = ""
    
    AdmissionDate6 = ""
    Day6 = ""
    Total6 = ""
    
    AdmissionDate7 = ""
    Day7 = ""
    Total7 = ""
    
    AdmissionDate8 = ""
    Day8 = ""
    Total8 = ""
    
    AdmissionDate9 = ""
    Day9 = ""
    Total9 = ""
    
    AdmissionDate10 = ""
    Day10 = ""
    Total10 = ""
    
    AdmissionDate11 = ""
    Day11 = ""
    Total11 = ""
    
    AdmissionDate12 = ""
    Day12 = ""
    Total12 = ""
    
    AdmissionDate13 = ""
    Day13 = ""
    Total13 = ""
    
    AdmissionDate14 = ""
    Day14 = ""
    Total14 = ""
    
    AdmissionDate15 = ""
    Day15 = ""
    Total15 = ""
    
    AdmissionDate16 = ""
    Day16 = ""
    Total16 = ""
    
    AdmissionDate17 = ""
    Day17 = ""
    Total17 = ""
    
    AdmissionDate18 = ""
    Day18 = ""
    Total18 = ""
    
    AdmissionDate19 = ""
    Day19 = ""
    Total19 = ""
    
    AdmissionDate20 = ""
    Day20 = ""
    Total20 = ""
    
    AdmissionDate21 = ""
    Day21 = ""
    Total21 = ""
    
    AdmissionDate22 = ""
    Day22 = ""
    Total22 = ""
    
    AdmissionDate23 = ""
    Day23 = ""
    Total23 = ""
    
    AdmissionDate24 = ""
    Day24 = ""
    Total24 = ""
    
    AdmissionDate25 = ""
    Day25 = ""
    Total25 = ""
    
    AdmissionDate26 = ""
    Day26 = ""
    Total26 = ""
    
    AdmissionDate27 = ""
    Day27 = ""
    Total27 = ""
    
    AdmissionDate28 = ""
    Day28 = ""
    Total28 = ""
    
    AdmissionDate29 = ""
    Day29 = ""
    Total29 = ""
    
    AdmissionDate30 = ""
    Day30 = ""
    Total30 = ""
    
    AdmissionDate31 = ""
    Day31 = ""
    Total31 = ""
    
End Function
'=================== Clear String Records ========================

' =================== Search Year Admission / Discharge Results ======================
Private Function SearchYearReports()
Const METHOD_NAME = "SearchYearReports"
On Error GoTo debugs
    If TxtYear = "" Then
        MsgBox "Please Enter Year", vbOKOnly + vbCritical, DialogBoxTitle
        TxtYear.SetFocus
    Else
        strAppend = ""
    
        If Len(Trim(TxtYear)) Then
            strAppend = strAppend + " AND Year = '" & Trim(TxtYear) & "'"
        End If
        
        If Option33.Value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option34.Value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option35.Value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option36.Value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option37.Value = True Then
            sql = ""
            sql = "AND Claimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option38.Value = True Then
            sql = ""
            sql = "AND Claimability = 'R'"
            strAppend = strAppend + sql
        End If

    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
' =================== Search Year Admission / Discharge Results ======================

'================ Admission / Discharge Year Listing ===========================
Private Sub YearListing(Optional strAppend As String)
Const METHOD_NAME = "YearListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
       
    Payor = ""
    ListView2.ListItems.Clear
    
    'sql = " SELECT Distinct PAYCode From VIEW_CLMReport Where 0 = 0 Order by PAYCode"
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "36"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                    
            With rst2
                If Len(rst2!PAYCode) Then
                
                Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Payor = rst2!PAYCode
                
                Call SearchVisitYearRecords(strAppend)
                
                If Len(Total1) Then
                    Record.SubItems(1) = Total1
                End If
                                
                If Len(Total2) Then
                    Record.SubItems(2) = Total2
                End If
                
                If Len(Total3) Then
                    Record.SubItems(3) = Total3
                End If
                
                If Len(Total4) Then
                    Record.SubItems(4) = Total4
                End If
                                
                If Len(Total5) Then
                    Record.SubItems(5) = Total5
                End If
                
                If Len(Total6) Then
                    Record.SubItems(6) = Total6
                End If
                
                If Len(Total7) Then
                    Record.SubItems(7) = Total7
                End If
                
                If Len(Total8) Then
                    Record.SubItems(8) = Total8
                End If
                
                If Len(Total9) Then
                    Record.SubItems(9) = Total9
                End If
                
                If Len(Total10) Then
                    Record.SubItems(10) = Total10
                End If
                
                If Len(Total11) Then
                    Record.SubItems(11) = Total11
                End If
                
                If Len(Total12) Then
                    Record.SubItems(12) = Total12
                End If
                End If
            End With
            rst2.MoveNext
        Loop
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
'================ Visit Year Listing ===========================

'==================== Search One Year Visit Records ===========================
Private Sub SearchVisitYearRecords(Optional strAppend As String)
Const METHOD_NAME = "SearchVisitYearRecords"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim strAppend1 As String

    ListView3.ListItems.Clear
    Call ClearBoolean
            
    strAppend1 = strAppend + " AND PAYCode = '" & Trim(Payor) & "'"
    'sql = " SELECT PAYCode, Month(DateVisit) As Month, " & _
          " COUNT(CASNumber) AS TotalCases " & _
          " FROM VIEW_CLMReport Where PAYCode = '" & Trim(Payor) & "'"
    
    'sql2 = " GROUP BY PAYCode, MONTH(DateVisit), YEAR(DateVisit) ORDER BY MONTH(DateVisit) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "97"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                               
                Set Record = ListView3.ListItems.Add(, , rst2!PAYCode)
    
                    If rst2!Month = "1" And one = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total1 = rst2!TotalCases
                            one = True
                        End If
                
                    ElseIf rst2!Month = "2" And two = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total2 = rst2!TotalCases
                            two = True
                        End If
                
                    ElseIf rst2!Month = "3" And three = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total3 = rst2!TotalCases
                            three = True
                        End If
                
                    ElseIf rst2!Month = "4" And four = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total4 = rst2!TotalCases
                            four = True
                        End If
                
                    ElseIf rst2!Month = "5" And five = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total5 = rst2!TotalCases
                            five = True
                        End If
                
                    ElseIf rst2!Month = "6" And six = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total6 = rst2!TotalCases
                            six = True
                        End If
                
                    ElseIf rst2!Month = "7" And seven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total7 = rst2!TotalCases
                            seven = True
                        End If
                
                    ElseIf rst2!Month = "8" And eight = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total8 = rst2!TotalCases
                            eight = True
                        End If
                
                    ElseIf rst2!Month = "9" And nine = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total9 = rst2!TotalCases
                            nine = True
                        End If
                
                    ElseIf rst2!Month = "10" And ten = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total10 = rst2!TotalCases
                            ten = True
                        End If
                
                    ElseIf rst2!Month = "11" And eleven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total11 = rst2!TotalCases
                            eleven = True
                        End If
                
                    ElseIf rst2!Month = "12" And twelve = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total12 = rst2!TotalCases
                            twelve = True
                        End If
                    End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'==================== Search One Year Visit Records ===========================

Private Function SearchRecords()
Const METHOD_NAME = "SearchRecords"
On Error GoTo debugs
Dim sql As String
Dim strAppend As String

    strAppend = ""
    GRPCompany = ""
    Clinics = ""
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend + " AND HLTCode ='" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
    End If

    If Len(Mid(CboInsType, 1, 1)) Then
        strAppend = strAppend & " AND INSCode = '" & Trim(Mid(CboInsType, 1, 1)) & "'"
    End If
    
    If ChkHLT.Value = 1 Then
        strAppend = strAppend & " AND (HLTCode Is null OR HLTCode = '')"
    End If
    
    If ChkINS.Value = 1 Then
        strAppend = strAppend & " AND (INSCode Is null OR INSCode = '')"
    End If
    
    If Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(cboSuppType) = "Both" Then
        strAppend = strAppend & " AND (MBMMemberType = 'P' Or MBMMemberType = 'Q') And CoverID = '00' "
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND MBMMemberType = 'P'  And CoverID = '00'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND MBMMemberType = 'Q' And CoverID = '00'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(cboSuppType) = "Both" Then
        strAppend = strAppend & " AND (MBMCMemberType = 'S' OR MBMCMemberType = 'C')"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND MBMCMemberType = 'C'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND MBMCMemberType = 'S'"
    ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND (MBMMemberType = 'P' OR MBMCMemberType = 'C')"
    ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND (MBMMemberType = 'Q' OR MBMCMemberType = 'S')"
    End If
    
    If Len(Mid(CboRace, 1, 1)) Then
        strAppend = strAppend & " AND RACCode = '" & Trim(Mid(CboRace, 1, 1)) & "'"
    End If
    
    If ChkRace.Value = 1 Then
        strAppend = strAppend & " AND (RACCode Is null OR RACCode = '')"
    End If
    
    If Len(Mid(CboSex, 1, 1)) Then
        strAppend = strAppend & " AND Sex = '" & Trim(Mid(CboSex, 1, 1)) & "'"
    End If
    
    If ChkSex.Value = 1 Then
        strAppend = strAppend & " AND (Sex Is null OR Sex = '')"
    End If
    
    If Len(Mid(CboRelation, 1, 1)) Then
        strAppend = strAppend & " AND Relation = '" & Trim(Mid(CboRelation, 1, 1)) & "'"
    End If
    
    If ChkRelation.Value = 1 Then
        strAppend = strAppend & " AND (Relation Is null OR Relation = '')"
    End If
       
    If Len(Mid(cboDataStatus, 1, 1)) Then
        strAppend = strAppend & " AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, 1)) & "'"
    End If
    
    If ChkDataStatus.Value = 1 Then
        strAppend = strAppend & " AND (MBMDataStatus Is null OR MBMDataStatus = '')"
    End If
    
    If Len(Mid(CboClaimability, 1, 1)) Then
        strAppend = strAppend & " AND Claimability = '" & Trim(Mid(CboClaimability, 1, 1)) & "'"
    End If
    
    If ChkClaimability.Value = 1 Then
        strAppend = strAppend & " AND (Claimability Is null OR Claimability = '')"
    End If
    
    If Len(Mid(CboStatus, 1, 1)) Then
        strAppend = strAppend & " AND MBMStatus = '" & Trim(Mid(CboStatus, 1, 1)) & "'"
    End If
    
    If ChkPolStatus.Value = 1 Then
        strAppend = strAppend & " AND (MBMStatus Is null OR MBMStatus = '')"
    End If
    
    If Mid(CboCaseStatus, 1, 1) = 0 Then
        strAppend = strAppend & " AND CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = 1 Then
        strAppend = strAppend & " AND (CASStatus = 'O' OR CASStatus = 'C')"
    ElseIf Mid(CboCaseStatus, 1, 1) = 2 Then
        strAppend = strAppend & " AND CASStatus = 'D'"
    End If
    
    If ChkCaseStatus.Value = 1 Then
        strAppend = strAppend & " AND (CASStatus Is null OR CASStatus = '')"
    End If
        
    If Len(Trim(Mid(CboExpired, 1, 1))) Then
          strAppend = strAppend & "AND MBMExpiredStatus = '" & Trim(Mid(CboExpired, 1, 1)) & "' "
    End If
    
    If ChkExpired.Value = 1 Then
        strAppend = strAppend & " AND (MBMExpiredStatus Is null OR MBMExpiredStatus = '')"
    End If
        
    If Len(Mid(CboPreMode, 1, 1)) Then
        strAppend = strAppend & " AND MBMInstallment = '" & Trim(Mid(CboPreMode, 1, 1)) & "'"
    End If
    
    If ChkPreMode.Value = 1 Then
        strAppend = strAppend & " AND (MBMInstallment Is null OR MBMInstallment = '')"
    End If
    
    If Len(Trim(CboState)) Then
          strAppend = strAppend & "AND STACode= '" & Trim(CboState) & "' "
    End If
    
    If ChkState.Value = 1 Then
        strAppend = strAppend & " AND (STACode Is null OR STACode = '')"
    End If
    
    If chkDept.Value = 1 Then
        strAppend = strAppend & " AND (MBMDepartment Is null OR MBMDepartment = '')"
    End If
    
    If Mid(cboRenewalInd, 1, 1) = "P" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'P')"
    ElseIf Mid(cboRenewalInd, 1, 1) = "N" Then
        strAppend = strAppend & " AND (MBMRenewFlag Is null OR MBMRenewFlag = '' OR MBMRenewFlag = 'N') "
    ElseIf Mid(cboRenewalInd, 1, 1) = "R" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'R')"
    ElseIf Mid(cboRenewalInd, 1, 1) = "T" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'T')"
    End If
    
    If ChkRenewal.Value = 1 Then
        strAppend = strAppend & " AND (MBMRenewFlag Is null OR MBMRenewFlag = '')"
    End If
    
    If Len(Mid(cboTakeOver, 1, 1)) Then
        strAppend = strAppend & " AND MBMTakeOver = '" & Trim(Mid(cboTakeOver, 1, 1)) & "'"
    End If
    
    If ChkTakeOver.Value = 1 Then
        strAppend = strAppend & " AND (MBMTakeOver Is null OR MBMTakeOver = '')"
    End If
    
    If Len(Trim(txtAgentCode)) Then
       strAppend = strAppend + " And  MBMAgentCode like '%" & ChkStr(txtAgentCode) & "%' "
    End If
    
    If ChkAgentCode.Value = 1 Then
        strAppend = strAppend & " AND (MBMAgentCode Is null OR MBMAgentCode = '')"
    End If
    
    If Len(Trim(txtPolNo)) Then
       strAppend = strAppend + " And  MBMPolicyNo like '%" & ChkStr(txtPolNo) & "%' "
    End If
    
    If ChkPolNo.Value = 1 Then
        strAppend = strAppend & " AND (MBMPolicyNo Is null OR MBMPolicyNo = '')"
    End If
    
    If Len(Trim(TxtPatientName)) Then
       strAppend = strAppend + " And  PatientName like '%" & ChkStr(TxtPatientName) & "%' "
    End If
    
    If Len(Trim(TxtAgeband1)) Then
        strAppend = strAppend + " AND AGEBand >= '" & Trim(TxtAgeband1) & "'"
    End If
    
    If Len(Trim(TxtAgeband2)) Then
        strAppend = strAppend + " AND AGEBand <= '" & Trim(TxtAgeband2) & "'"
    End If
    
    If Len(Trim(txtPOLYear1)) Then
        strAppend = strAppend & " AND MBMPolicyYear >= '" & Trim(txtPOLYear1) & "'"
    End If
    
    If Len(Trim(txtPOLYear2)) Then
        strAppend = strAppend & " AND MBMPolicyYear <= '" & Trim(txtPOLYear2) & "'"
    End If
    
    If ChkPolYear.Value = 1 Then
        strAppend = strAppend & " AND (MBMPolicyYear Is null OR MBMPolicyYear = '')"
    End If
    
    If Len(Trim(TxtMCDay1)) Then
        strAppend = strAppend & " AND MCDays >= '" & Trim(TxtMCDay1) & "'"
    End If
    
    If Len(Trim(TxtMCDay2)) Then
        strAppend = strAppend & " AND MCDays <= '" & Trim(TxtMCDay2) & "'"
    End If
    
    If ChkMCDay.Value = 1 Then
        strAppend = strAppend & " AND (MCDays Is null OR MCDays = '')"
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Trim(Mid(cboPayorCode, 1, 2)) & "'"
    End If
           
    If ChkPayor.Value = 1 Then
        strAppend = strAppend & " AND (PAYCode Is null OR PAYCode = '')"
    End If
     
    If Len(Trim(CboGrpCom)) Then
         GRPCompany = Replace(CboGrpCom, "'", "''")
         strAppend = strAppend + " AND MBMGroupCompany = '" & Trim(GRPCompany) & "'"
    End If
    
    If ChkGrpCom.Value = 1 Then
        strAppend = strAppend & " AND (MBMGroupCompany Is null OR MBMGroupCompany = '')"
    End If
    
    If Len(Trim(cboCLNName)) Then
        strAppend = strAppend + " And ClinicCode = '" & Trim(Mid(cboCLNName, 1, IIf(InStr(1, cboCLNName, "-") = 0, 4, InStr(1, cboCLNName, "-") - 1))) & "'"
    End If
    
    If ChkHOS.Value = 1 Then
        strAppend = strAppend & " AND (CLNName Is null OR CLNName = '')"
    End If
    
    If Len(Trim(cboCLNGrp)) Then
        strAppend = strAppend + " And ClinicGRPName = '" & Trim(cboCLNGrp) & "'"
    End If
    
    If Len(Trim(TxtBranch)) Then
       strAppend = strAppend + " And  BRNCode = '" & ChkStr(TxtBranch) & "'"
    End If
    
    If Len(Trim(TxtDOV1)) > 0 And IsDate(TxtDOV1) = True Then
        strAppend = strAppend + " And DateVisit >= '" & Format(TxtDOV1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOV2)) > 0 And IsDate(TxtDOV2) = True Then
        strAppend = strAppend + " And DateVisit <= '" & Format(TxtDOV2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkDOA.Value = 1 Then
        strAppend = strAppend & " AND (DateVisit Is null OR DateVisit = '')"
    End If
    
    If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And RegisterDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And RegisterDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
    End If
    
    If chkDOR.Value = 1 Then
        strAppend = strAppend & " AND (RegisterDate Is null OR RegisterDate = '')"
    End If
    
    If Len(Trim(txtDOW1)) > 0 And IsDate(txtDOW1) = True Then
        strAppend = strAppend + " And ClaimOpenDate >= '" & Format(txtDOW1, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtDOW2)) > 0 And IsDate(txtDOW2) = True Then
        strAppend = strAppend + " And ClaimOpenDate <= '" & Format(txtDOW2, "mm/dd/yyyy") & "'"
    End If

    If ChkDOW.Value = 1 Then
        strAppend = strAppend & " AND (ClaimOpenDate Is null OR ClaimOpenDate = '')"
    End If
    
    If Len(Trim(TxtInvDate1)) > 0 And IsDate(TxtInvDate1) = True Then
        strAppend = strAppend + " And InvoiceDate >= '" & Format(TxtInvDate1, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(TxtInvDate2)) > 0 And IsDate(TxtInvDate2) = True Then
        strAppend = strAppend + " And InvoiceDate <= '" & Format(TxtInvDate2, "mm/dd/yyyy") & "'"
    End If

    If ChkInvoiceDate.Value = 1 Then
        strAppend = strAppend & " AND (InvoiceDate Is null OR InvoiceDate = '')"
    End If
    
    If Len(Trim(txtTreatment1)) Then
        strAppend = strAppend + " And CASTreatReason1 >= '" & Trim(txtTreatment1) & "'"
    End If
    
    If Len(Trim(txtTreatment2)) Then
        strAppend = strAppend + " And CASTreatReason1 <= '" & Trim(txtTreatment2) & "'"
    End If
    
    If ChkFirst.Value = 1 Then
        strAppend = strAppend & " AND (CASTreatReason1 Is null OR CASTreatReason1 = '')"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend & " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend & " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If ChkPlan.Value = 1 Then
        strAppend = strAppend & " AND (PLNCode Is null OR PLNCode = '')"
    End If
    
    If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtPayEffDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkEffDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMPayorEffDate Is null OR MBMPayorEffDate = '')"
    End If
    
    If Len(Trim(TxtPayExpDate1)) And IsDate(TxtPayExpDate1) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate >= '" & Format(TxtPayExpDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayExpDate2)) And IsDate(TxtPayExpDate2) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate <= '" & Format(TxtPayExpDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkExpDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMPayorExpDate Is null OR MBMPayorExpDate = '')"
    End If
    
    If Len(Trim(txtDateJoin1)) And IsDate(txtDateJoin1) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate >= '" & Format(txtDateJoin1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateJoin2)) And IsDate(txtDateJoin2) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate <= '" & Format(txtDateJoin2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkJoinDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMFirstJoinedDate Is null OR MBMFirstJoinedDate = '')"
    End If
    
    If Len(Trim(txtDOB1)) And IsDate(txtDOB1) = True Then
        strAppend = strAppend & " AND DOB >= '" & Format(txtDOB1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOB2)) And IsDate(txtDOB2) = True Then
        strAppend = strAppend & " AND DOB <= '" & Format(txtDOB2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkDOB.Value = 1 Then
        strAppend = strAppend & " AND (DOB Is null OR DOB = '')"
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
       strAppend = strAppend + " AND MBMNumber >= '" & RemStr(Trim(TxtMBMNo1)) & "'"
    End If
    
    If Len(Trim(TxtMBMNo2)) Then
       strAppend = strAppend + " AND MBMNumber <= '" & RemStr(Trim(TxtMBMNo2)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo1)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
        strAppend = strAppend & " AND BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
        strAppend = strAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
    End If
    
    If ChkBordx.Value = 1 Then
        strAppend = strAppend & " AND (BordxRefNo Is null OR BordxRefNo = '')"
    End If
    
    If Len(Trim(TxtBordxDate1)) And IsDate(TxtBordxDate1) = True Then
        strAppend = strAppend & " AND BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxDate2)) And IsDate(TxtBordxDate2) = True Then
        strAppend = strAppend & " AND BordxDate <= '" & Format(TxtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkBordxDate.Value = 1 Then
        strAppend = strAppend & " AND (BordxDate Is null OR BordxDate = '')"
    End If
    
    If Len(Trim(TxtICNo1)) Then
        strAppend = strAppend + " AND ICNo >= '" & Trim(TxtICNo1) & "'"
    End If
    
    If Len(Trim(TxtICNo2)) Then
        strAppend = strAppend + " AND ICNo <= '" & Trim(TxtICNo2) & "'"
    End If
    
    If ChkICNo.Value = 1 Then
        strAppend = strAppend + " AND (ICNo Is null or ICNo = '')"
    End If
    
    If Len(Trim(TxtMBMIc1)) Then
        strAppend = strAppend + " AND MBMIcBcPp >= '" & Trim(TxtMBMIc1) & "'"
    End If
    If Len(Trim(TxtMBMIc2)) Then
        strAppend = strAppend + " AND MBMIcBcPp <= '" & Trim(TxtMBMIc2) & "'"
    End If
    If ChkMBMIC.Value = 1 Then
        strAppend = strAppend + " AND (MBMIcBcPp Is null OR MBMIcBcPp = '')"
    End If
    
    If Len(txtActualAmt1) Then
        strAppend = strAppend + " AND CLMTotalActual >= " & Trim(txtActualAmt1) & ""
    End If
        
    If Len(txtActualAmt2) Then
        strAppend = strAppend + " AND CLMTotalActual <= " & Trim(txtActualAmt2) & ""
    End If
    
    If Len(txtAppAmt1) Then
        strAppend = strAppend + " AND CLMTotalApproved >= " & Trim(txtAppAmt1) & ""
    End If
        
    If Len(txtAppAmt2) Then
        strAppend = strAppend + " AND CLMTotalApproved <= " & Trim(txtAppAmt2) & ""
    End If
    
    If Len(TxtOSUtilized1) Then
        strAppend = strAppend + " AND OSUtilized >= " & Trim(TxtOSUtilized1) & ""
    End If
        
    If Len(TxtOSUtilized2) Then
        strAppend = strAppend + " AND OSUtilized <= " & Trim(TxtOSUtilized2) & ""
    End If
    
    If Mid(cboPanelStatus, 1, 1) = "C" Then
        strAppend = strAppend + " AND CLNPanel = 'C'"
    ElseIf Mid(cboPanelStatus, 1, 1) = "Y" Then
        strAppend = strAppend + " AND CLNPanel = 'Y'"
    End If
    
    If Mid(CboClmPending, 1, 1) = "N" Then
        strAppend = strAppend + " AND (CLMPending = 'N' OR CLMPending = '' OR CLMPending IS NULL)"
    ElseIf Mid(CboClmPending, 1, 1) = "Y" Then
        strAppend = strAppend + " AND CLMPending = 'Y'"
    End If
        
    If Len(Trim(TxtYear)) Then
        strAppend = strAppend + " AND DOVYear = '" & Trim(TxtYear) & "'"
    End If
    
    If Len(Trim(TxtInvoice1)) Then
        strAppend = strAppend + " AND InvoiceNo >= '" & Trim(TxtInvoice1) & "'"
    End If
    
    If Len(Trim(TxtInvoice2)) Then
        strAppend = strAppend + " AND InvoiceNo <= '" & Trim(TxtInvoice2) & "'"
    End If
    
    If ChkInvoice.Value = 1 Then
        strAppend = strAppend + " AND (InvoiceNo Is null OR InvoiceNo = '')"
    End If
    
    If Len(Trim(TxtSubmission1)) Then
        strAppend = strAppend + " AND BillNo >= '" & Trim(TxtSubmission1) & "'"
    End If
    
    If Len(Trim(TxtSubmission2)) Then
        strAppend = strAppend + " AND BillNo <= '" & Trim(TxtSubmission2) & "'"
    End If
    
    If ChkSubmission.Value = 1 Then
        strAppend = strAppend + " AND (BillNo Is null OR BillNo = '')"
    End If
    
    If Len(Trim(TxtOPBatchNo1)) Then
        strAppend = strAppend + " AND OPBatchNo >= '" & Trim(TxtOPBatchNo1) & "'"
    End If
    
    If Len(Trim(TxtOPBatchNo2)) Then
        strAppend = strAppend + " AND OPBatchNo <= '" & Trim(TxtOPBatchNo2) & "'"
    End If
    
    If Len(Trim(txtCLNCode1)) Then
        strAppend = strAppend + " AND ClinicCode >= '" & Trim(txtCLNCode1) & "'"
    End If
    
    If Len(Trim(txtCLNCode2)) Then
        strAppend = strAppend + " AND ClinicCode <= '" & Trim(txtCLNCode2) & "'"
    End If
    
    If Mid(CboClinicType, 1, 1) = 0 Then
        strAppend = strAppend & " AND CLNType = 'S'"
    ElseIf Mid(CboClinicType, 1, 1) = 1 Then
        strAppend = strAppend & " AND CLNType = 'G'"
    ElseIf Mid(CboClinicType, 1, 1) = 2 Then
        strAppend = strAppend & " AND (CLNType = '' OR CLNType IS NULL)"
    End If
    
    If Len(Mid(CboPayType, 1, 1)) Then
        strAppend = strAppend & " AND INVPayTo = '" & Trim(Mid(CboPayType, 1, 1)) & "'"
    End If
    
    If ChkAmt.Value = 1 Then
        strAppend = strAppend & " AND VSAmt <> 0"
    End If
    
    If cboDeptSelection <> "" Then
        strAppend = strAppend & " AND MBMDepartment = '" & Trim(cboDeptSelection) & "'"
    End If
    
    If CboIEinv = "Yes" Then
        strAppend = strAppend + " AND substring(InvoiceNo,1,2)='IE' "
    End If
    
    If CboReportType.ListIndex = 0 And CboSelection.Text = "Daily Report" Then
        Call DailyVisitListing(strAppend)
    ElseIf CboReportType.ListIndex = 0 And CboSelection.Text = "Monthly Report" Then
        If Mid(TxtDOV1, 4, 2) <> Mid(TxtDOV2, 4, 2) Or Mid(TxtDOV1, 7, 4) <> Mid(TxtDOV2, 7, 4) Then
            MsgBox "Must key in same Month", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf TxtDOV1 = "__/__/____" Or TxtDOV2 = "__/__/____" Then
            MsgBox "Date of Visit can't be empty", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call MonthyReportListing(strAppend)
        End If
    ElseIf CboReportType.ListIndex = 0 And CboSelection.Text = "YTD Report" Then
        If TxtYear = "" Then
            MsgBox "Please enter Year", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call YearListing(strAppend)
        End If
        
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "Summary" Then
        Call AgeGroupListing(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "By Gender" Then
        Call AgeGenderListing(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "By Insured Type" Then
        Call AgeInsuredListing(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "By Race" Then
        Call AgeRaceListing(strAppend)
    ElseIf CboReportType.ListIndex = 1 And CboSelection.Text = "By Plan Code" Then
        Call AgePlanListing(strAppend)
        
    ElseIf CboReportType.ListIndex = 2 And CboSelection.Text = "Details Listing" Then
        Call AgentListing(strAppend)
    ElseIf CboReportType.ListIndex = 2 And CboSelection.Text = "General Summary" Then
        Call AgentSumListing(strAppend)
    ElseIf CboReportType.ListIndex = 2 And CboSelection.Text = "By Top 50 Claims" Then
        Call AgentSumListing(strAppend)
        
    ElseIf CboReportType.ListIndex = 3 And CboSelection.Text = "Summary Bordx" Then
        Call BordxNoListing(strAppend)
    ElseIf CboReportType.ListIndex = 3 And CboSelection.Text = "Bordx (By Grp Company)" Then
        Call BordxNoGrpListing(strAppend)
    ElseIf CboReportType.ListIndex = 3 And CboSelection.Text = "Bordx (By Grp Company & Department)" Then
        Call BordxNoDeptListing(strAppend)
        
    ElseIf (CboReportType.ListIndex = 4 And CboSelection.Text = "Top 20 by Claims") Or _
        (CboReportType.ListIndex = 4 And CboSelection.Text = "Top 20 by Amount") Then
        Call CLNTop20(strAppend)
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "General Summary (By Payor)" Then
        Call CLNSumListing(strAppend)
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "General Summary (All)" Then
        Call CLNSumListing2(strAppend)
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "Detail Listing" Then
        Call CLNDetailListing(strAppend)
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "By Treatment Reason" Then
        Call CLNByDiag(strAppend)
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "By Clinic Type" Then
        Call CLNByType(strAppend)
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "By Clinic Type & Clinic Panel" Then
        Call CLNByTypePanel(strAppend)
    ElseIf CboReportType.ListIndex = 4 And CboSelection.Text = "General Summary (By Grp Company)" Then
        Call CLNSumListing3(strAppend)
        
    ElseIf CboReportType.ListIndex = 5 Then
        Call ClinicGrpSumListing(strAppend)
    
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Summary (By Visit)" Then
        Call DepartmentByVisit(strAppend)
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Summary (By MC)" Then
        Call DepartmentByMC(strAppend)
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Employee Utilization" Then
        Call DepartmentByEmp(strAppend)
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Illness Utilization" Then
        Call DepartmentByIllness(strAppend)
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Clinic Utilization" Then
        Call DepartmentByClinic(strAppend)
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Department Utilization" Then
        Call DepartmentByDept(strAppend)
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "MC Times Analysis" Then
        Call DepartmentByMCTimes(strAppend)
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "MC Days Analysis" Then
        Call DepartmentByMCDays(strAppend)
    ElseIf CboReportType.ListIndex = 6 And CboSelection.Text = "Visit Analysis" Then
        Call DepartmentByVisit1(strAppend)
        
    ElseIf CboReportType.ListIndex = 7 And CboSelection.Text = "Top 20 by Claims" Then
        Call FinalDiagListing(strAppend)
    ElseIf CboReportType.ListIndex = 7 And CboSelection.Text = "Top 20 by Amount" Then
        Call FinalDiagListing(strAppend)
    ElseIf CboReportType.ListIndex = 7 And CboSelection.Text = "General Summary" Then
        Call FinalDiagListing(strAppend)
    ElseIf CboReportType.ListIndex = 7 And CboSelection.Text = "By Clinic" Then
        Call FinalDiagbyHos(strAppend)
    
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "Summary" Then
        Call SexListing(strAppend)
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "By Ageband" Then
        Call SexAgebandListing(strAppend)
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "By Insured Type" Then
        Call SexInsuredListing(strAppend)
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "By Race" Then
        Call SexRaceListing(strAppend)
    ElseIf CboReportType.ListIndex = 8 And CboSelection.Text = "By Plan Code" Then
        Call SexPlanListing(strAppend)
    
    ElseIf CboReportType.ListIndex = 9 And CboSelection.Text = "Without Remarks" Then
        Call ResultListing(strAppend)
    ElseIf CboReportType.ListIndex = 9 And CboSelection.Text = "With Remarks" Then
        Call ResultListing2(strAppend)
    ElseIf CboReportType.ListIndex = 9 And CboSelection.Text = "By No of Visits" Then
        Call ResultListing3(strAppend)
        
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "General Summary (By Payor)" Then
        Call GRPCOMSumListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "General Summary (By Grp Company)" Then
        Call GRPCOMSumListing2(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "Detail Listing" Then
        Call GRPCOMDetailListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "Employee Utilization" Then
        Call GRPCOMEmpUtilizeListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "Illness Utilization" Then
        Call GRPCOMIllnessListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "Clinic Utilization" Then
        Call GRPCOMClinicListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "Department Utilization" Then
        Call GRPCOMDeptListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "MC Times Analysis" Then
        Call GRPCOMMCTimesListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "MC Days Analysis" Then
        Call GRPCOMMCDaysListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "Visit Analysis" Then
        Call GRPCOMVisitListing(strAppend)
    ElseIf CboReportType.ListIndex = 10 And CboSelection.Text = "Group Company" Then
        Call GRPCOMGrpCOMListing(strAppend)
    ElseIf CboReportType.ListIndex = 11 Then
        Call ICSummary(strAppend)
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "Summary" Then
        Call INSListing(strAppend)
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "By Plan Code" Then
        Call PlanInsuredListing(strAppend)
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "By Ageband" Then
        Call AgeInsuredListing(strAppend)
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "By Gender" Then
        Call SexInsuredListing(strAppend)
    ElseIf CboReportType.ListIndex = 12 And CboSelection.Text = "By Race" Then
        Call RaceInsuredListing(strAppend)
    ElseIf CboReportType.ListIndex = 13 And CboSelection.Text = "Invoice Listing" Then
        Call InvoiceListing(strAppend)
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Summary" Then
        Call MCSummary(strAppend)
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Summary (By Clinic)" Then
        Call MCSummary2(strAppend)
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Summary (By Grp Company)" Then
        Call MCSummary3(strAppend)
    ElseIf CboReportType.ListIndex = 14 And CboSelection.Text = "Summary (By Diag.)" Then
        Call MCSummary4(strAppend)
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Eff. Month (By Payor)" Then
        Call EffMthSummary(strAppend)
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Eff. Month (All Payor)" Then
        Call EffMthSummary2(strAppend)
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Exp. Month (By Payor)" Then
        Call ExpMthSummary(strAppend)
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Exp. Month (All Payor)" Then
        Call ExpMthSummary2(strAppend)
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Visit Month (By Payor)" Then
        Call VisitMthSummary(strAppend)
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Visit Month (All Payor)" Then
        Call VisitMthSummary2(strAppend)
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Reg. Month (By Payor)" Then
        Call RegMthSummary(strAppend)
    ElseIf CboReportType.ListIndex = 15 And CboSelection.Text = "Reg. Month (All Payor)" Then
        Call RegMthSummary2(strAppend)
    ElseIf CboReportType.ListIndex = 16 And CboSelection.Text = "Utilization (By Member)" Then
        Call ReportMBMGeneralListing(strAppend)
    ElseIf CboReportType.ListIndex = 16 And CboSelection.Text = "Claims Listing" Then
        Call ReportCLMGeneralListing(strAppend)
    ElseIf CboReportType.ListIndex = 16 And CboSelection.Text = "Summary (by Clinics)" Then
        Call ReportCLNSummary(strAppend)
    ElseIf CboReportType.ListIndex = 16 And CboSelection.Text = "Summary (by Illness)" Then
        Call ReportIllnessSummary(strAppend)
    ElseIf CboReportType.ListIndex = 16 And CboSelection.Text = "Departmental Bordx" Then
        Call ReportDepartmentalBordx(strAppend)
    ElseIf CboReportType.ListIndex = 16 And CboSelection.Text = "O/S Utilized (Combined)" Then
        Call ReportMBMSummaryOS(strAppend)
    ElseIf CboReportType.ListIndex = 16 And CboSelection.Text = "O/S Utilized (Separate)" Then
        Call ReportMBMSummaryOS1(strAppend)
    ElseIf CboReportType.ListIndex = 17 Then
        Call PayorListing(strAppend)
    ElseIf CboReportType.ListIndex = 18 And CboSelection.Text = "General Summary with Premium" Then
        Call PlanPremiumListing(strAppend)
    ElseIf CboReportType.ListIndex = 18 And CboSelection.Text = "General Summary" Then
        Call PlanListing(strAppend)
    ElseIf CboReportType.ListIndex = 18 And CboSelection.Text = "Top 10 Clinics" Then
        Call PlanTopHosListing(strAppend)
    ElseIf CboReportType.ListIndex = 18 And CboSelection.Text = "Top 10 Treatment Reasons" Then
        Call PlanTopDiagListing(strAppend)
    ElseIf CboReportType.ListIndex = 18 And CboSelection.Text = "By Ageband" Then
        Call PlanAgebandListing(strAppend)
    ElseIf CboReportType.ListIndex = 18 And CboSelection.Text = "By Insured Type" Then
        Call PlanInsuredListing(strAppend)
    ElseIf CboReportType.ListIndex = 18 And CboSelection.Text = "By Insured Type & Ageband" Then
        Call PlanInsuredListing2(strAppend)
    ElseIf CboReportType.ListIndex = 18 And CboSelection.Text = "By Clinic Panel" Then
        Call PlanByPanel(strAppend)
    ElseIf CboReportType.ListIndex = 19 Then
        Call PolGRPCOMListing(strAppend)
    ElseIf CboReportType.ListIndex = 20 And CboSelection.Text = "Summary" Then
        Call RaceListing(strAppend)
    ElseIf CboReportType.ListIndex = 20 And CboSelection.Text = "By Ageband" Then
        Call RaceAgebandListing(strAppend)
    ElseIf CboReportType.ListIndex = 20 And CboSelection.Text = "By Insured Type" Then
        Call RaceInsuredListing(strAppend)
    ElseIf CboReportType.ListIndex = 20 And CboSelection.Text = "By Gender" Then
        Call RaceSexListing(strAppend)
    ElseIf CboReportType.ListIndex = 20 And CboSelection.Text = "By Plan Code" Then
        Call RacePlanListing(strAppend)
    ElseIf CboReportType.ListIndex = 21 Then
        Call StateListing(strAppend)
    ElseIf CboReportType.ListIndex = 22 And CboSelection.Text = "Summary (By Payor)" Then
        Call MemberSummaryByPayor(strAppend)
    ElseIf CboReportType.ListIndex = 22 And CboSelection.Text = "Summary (All)" Then
        Call MemberSummary(strAppend)
    ElseIf CboReportType.ListIndex = 22 And CboSelection.Text = "Summary (By Grp Company)" Then
        Call MemberSummaryByGrpCompany(strAppend)
    ElseIf CboReportType.ListIndex = 22 And CboSelection.Text = "O/S Utilized" Then
        Call MemberSummaryOS(strAppend)
    ElseIf CboReportType.ListIndex = 22 And CboSelection.Text = "Summary (By Illness)" Then
        Call UtilizationByIllness(strAppend)
    ElseIf CboReportType.ListIndex = 22 And CboSelection.Text = "Summary (By Member)" Then
        Call UtilizationByMBM(strAppend)
    ElseIf CboReportType.ListIndex = 23 And CboSelection.Text = "Claims Same Date" Then
        Call ClaimsSameDate(strAppend)
    ElseIf CboReportType.ListIndex = 23 And CboSelection.Text = "Claims Same Amount" Then
        Call ClaimsSameAmt(strAppend)
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub CLNTop20(Optional strAppend As String)
Const METHOD_NAME = "CLNTop20"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT TOP 20 PAYCode, ClinicCode, CLNName, COUNT(CLNName) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved" & _
          " From VIEW_CLMReport Where  0 = 0"
    
    If CboSelection.Text = "Top 20 by Claims" Then
        strAppendCase = "29"
    '    sql2 = " GROUP BY PAYCode, ClinicCode, CLNName ORDER BY Total DESC "
    ElseIf CboSelection.Text = "Top 20 by Amount" Then
        strAppendCase = "30"
    '    sql2 = " GROUP BY PAYCode, ClinicCode, CLNName ORDER BY TotalActual DESC "
    End If
    
    'sql = sql + strAppend + sql2
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
    
            With rst2
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!ClinicCode) Then
                    Record.SubItems(1) = Trim(!ClinicCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(2) = Trim(!CLNName)
                Else
                    Record.SubItems(2) = "Empty"
                End If
                                
                If Len(!total) Then
                    Record.SubItems(3) = !total
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                            
                If CDbl(!total) <> 0 Then
                    Record.SubItems(6) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(7) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                    Record.SubItems(7) = "0.00"
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
        
    If ListView2.ListItems.Count = 0 Then
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

Private Sub CLNSumListing(Optional strAppend As String)
Const METHOD_NAME = "CLNSumListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, ClinicCode, CLNName, COUNT(CLNName) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY PAYCode, ClinicCode, CLNName "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "26"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!ClinicCode) Then
                    Record.SubItems(1) = Trim(!ClinicCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(2) = Trim(!CLNName)
                Else
                    Record.SubItems(2) = "Empty"
                End If
                                
                If Len(!total) Then
                    Record.SubItems(3) = !total
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If CDbl(!total) <> 0 Then
                    Record.SubItems(6) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(7) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                    Record.SubItems(7) = "0.00"
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub CLNSumListing2(Optional strAppend As String)
Const METHOD_NAME = "CLNSumListing2"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT ClinicCode, CLNName, CLNType, CLNLocation, SUM(MCDays) AS MCDays, COUNT(CLNName) AS Total, " & _
          " SUM(MCStatus) AS MCTimes, SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY ClinicCode, CLNName, CLNType, CLNLocation "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "27"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!ClinicCode) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!ClinicCode))
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(1) = Trim(!CLNName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!CLNType) Then
                    Record.SubItems(2) = Trim(!CLNType)
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!CLNLocation) Then
                    Record.SubItems(3) = Trim(!CLNLocation)
                Else
                    Record.SubItems(3) = "Empty"
                End If
                
                If Len(!total) Then
                    Record.SubItems(4) = !total
                End If
                
                If Len(!MCTimes) Then
                    Record.SubItems(5) = !MCTimes
                Else
                    Record.SubItems(5) = "0"
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(6) = !MCDays
                End If
                                
                If Len(!TotalActual) Then
                   Record.SubItems(7) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(8) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If CDbl(!total) <> 0 Then
                    Record.SubItems(9) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(10) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                Else
                    Record.SubItems(9) = "0.00"
                    Record.SubItems(10) = "0.00"
                End If
                
                If !MCTimes <> "" Then
                    If CDbl(!MCTimes) <> 0 Then
                        Record.SubItems(11) = Round(!MCTimes / !total * 100, 2)
                        Record.SubItems(12) = Round(!MCDays / !MCTimes, 2)
                    Else
                        Record.SubItems(11) = "0"
                        Record.SubItems(12) = "0"
                    End If
                Else
                    Record.SubItems(11) = "0"
                    Record.SubItems(12) = "0"
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub CLNSumListing3(Optional strAppend As String)
Const METHOD_NAME = "CLNSumListing3"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMGroupCompany, ClinicCode, CLNName, COUNT(CLNName) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY MBMGroupCompany, ClinicCode, CLNName "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "28"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!ClinicCode) Then
                    Record.SubItems(1) = Trim(!ClinicCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(2) = Trim(!CLNName)
                Else
                    Record.SubItems(2) = "Empty"
                End If
                                
                If Len(!total) Then
                    Record.SubItems(3) = !total
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If CDbl(!total) <> 0 Then
                    Record.SubItems(6) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(7) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                    Record.SubItems(7) = "0.00"
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
        
    If ListView2.ListItems.Count = 0 Then
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

Private Sub CLNByDiag(Optional strAppend As String)
Const METHOD_NAME = "CLNByDiag"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim StrSql As String

    lblCurrent = 0
    ListView2.ListItems.Clear
    DIagCode = ""
    Description = ""
    
    'sql = " SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual, ClinicCode, CLNName " & _
          " From VIEW_CLMReport Where 0 = 0 " & strAppend & _
          " GROUP BY CASTreatReason1, ClinicCode, CLNName "
    
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "23"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
       
        Do Until rst1.EOF
    
            With rst1
                               
                If Len(!ClinicCode) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!ClinicCode))
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                If Len(!CLNName) Then
                    Record.SubItems(1) = Trim(!CLNName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                If Len(!CASTreatReason1) Then
                    Record.SubItems(2) = Trim(!CASTreatReason1)
                    DIagCode = Trim(!CASTreatReason1)
                Else
                    Record.SubItems(2) = "Empty"
                    DIagCode = ""
                End If
                
                'sql1 = " SELECT ReasonCode, ReasonDesc From TreatmentReason Where ReasonCode = '" & Trim(DIagCode) & "'"
                'rst2.Open sql1, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                StrSql = "EXEC CISClaimTracking '" & Trim(DIagCode) & "',10"
                rst2.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                       
                If rst2.EOF = True Then
                    rst2.Close
                    Set rst2 = Nothing
                    Record.SubItems(3) = ""
                Else
                    Description = Trim(rst2!ReasonDesc)
                    Record.SubItems(3) = Trim(Description)
                    rst2.Close
                    Set rst2 = Nothing
                End If
                
                If Len(!total) Then
                    Record.SubItems(4) = !total
                End If
                                
                If Len(!TotalActual) Then
                   Record.SubItems(5) = Format(!TotalActual, "##,#0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(6) = Format(!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst1.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CLNByType(Optional strAppend As String)
Const METHOD_NAME = "CLNByType"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
    DIagCode = ""
    Description = ""
        
    'sql = " SELECT COUNT(CASNumber) AS Total, CLNType, Sum(CLMTotalApproved) As TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 " & strAppend & _
          " GROUP BY CLNType "
    
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "24"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
       
        Do Until rst1.EOF
    
            With rst1
                               
                If Len(!CLNType) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!CLNType))
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!total) Then
                    Record.SubItems(1) = !total
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(2) = Format(!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst1.MoveNext
        Loop
        
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CLNByTypePanel(Optional strAppend As String)
Const METHOD_NAME = "CLNByTypePanel"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim strAppend1 As String

    lblCurrent = 0
    ListView2.ListItems.Clear
    ClinicType = ""
        
    'sql = " SELECT CLNType, Count(CASNUmber) AS TotalClaim, Sum(CLMTotalApproved) As TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 " & strAppend & _
          " GROUP BY CLNType "
    
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "24"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
       
        Do Until rst1.EOF
    
            With rst1
                               
                If Len(!CLNType) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!CLNType))
                    ClinicType = Trim(!CLNType)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    ClinicType = ""
                End If
                
                If ClinicType = "" Then
                
                    strAppendCase = "31"
                    strAppend1 = strAppend + " AND (CLNType IS NULL OR CLNType = '') AND CLNPanel = 'C'"
                    Call CalCLMAmt(strAppend1, strAppendCase, Record)
                    
                    'sql2 = " SELECT CLNType, Count(CASNumber) AS TotalClaimP, SUM(CLMTotalApproved) AS TotalPanelAmt " & _
                           " From VIEW_CLMReport Where (CLNType IS NULL OR CLNType = '') AND CLNPanel = 'C'"
              
                    'sql3 = " GROUP BY CLNType "
                
                    'If Len(Trim(strAppend)) Then
                    '   sql2 = sql2 + strAppend + sql3
                    'Else
                    '   sql2 = sql2 + sql13
                    'End If
        
                    'rst2.Open sql2, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                    
                    'Record.SubItems(1) = IIf(IsNumeric(rst2!TotalClaimP) = True, rst2!TotalClaimP, 0)
                    'Record.SubItems(2) = Format(IIf(IsNumeric(rst2!TotalPanelAmt) = True, rst2!TotalPanelAmt, 0), "#,##0.00")
                
                    'rst2.Close
                    'Set rst2 = Nothing
                    
                    strAppendCase = "32"
                    strAppend1 = strAppend + " AND (CLNType IS NULL OR CLNType = '') AND (CLNPanel <> 'C' OR CLNPanel IS NULL OR CLNPanel = '')"
                    Call CalCLMAmt(strAppend1, strAppendCase, Record)
                    
                    'sql4 = " SELECT CLNType, Count(CASNumber) AS TotalClaimNP, SUM(CLMTotalApproved) AS TotalNonPanelAmt " & _
                          " From VIEW_CLMReport Where (CLNType IS NULL OR CLNType = '') AND (CLNPanel <> 'C' OR CLNPanel IS NULL OR CLNPanel = '')"
                  
                    'sql5 = " GROUP BY CLNType "
                    
                    'If Len(Trim(strAppend)) Then
                    '    sql4 = sql4 + strAppend + sql5
                    'Else
                    '    sql4 = sql4 + sql5
                    'End If
            
                    'rst3.Open sql4, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                                 
                    'Record.SubItems(3) = IIf(IsNumeric(rst3!TotalClaimNP) = True, rst3!TotalClaimNP, 0)
                    'Record.SubItems(4) = Format(IIf(IsNumeric(rst3!TotalNonPanelAmt) = True, rst3!TotalNonPanelAmt, 0), "#,##0.00")
                    
                    'rst3.Close
                    'Set rst3 = Nothing
                    
                Else
                    
                    strAppendCase = "31"
                    strAppend1 = strAppend + " AND CLNType = '" & Trim(ClinicType) & "' AND CLNPanel = 'C'"
                    Call CalCLMAmt(strAppend1, strAppendCase, Record)
                    
                    'sql2 = " SELECT CLNType, Count(CASNumber) AS TotalClaimP, SUM(CLMTotalApproved) AS TotalPanelAmt " & _
                           " From VIEW_CLMReport Where CLNType = '" & Trim(ClinicType) & "' AND CLNPanel = 'C'"
              
                    'sql3 = " GROUP BY CLNType "
                
                    'If Len(Trim(strAppend)) Then
                    '   sql2 = sql2 + strAppend + sql3
                    'Else
                    '   sql2 = sql2 + sql13
                    'End If
        
                    'rst2.Open sql2, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                                  
                    'Record.SubItems(1) = IIf(IsNumeric(rst2!TotalClaimP) = True, rst2!TotalClaimP, 0)
                    'Record.SubItems(2) = Format(IIf(IsNumeric(rst2!TotalPanelAmt) = True, rst2!TotalPanelAmt, 0), "#,##0.00")
                    
                    'rst2.Close
                    'Set rst2 = Nothing
                    
                    strAppendCase = "32"
                    strAppend1 = strAppend + " AND CLNType = '" & Trim(ClinicType) & "' AND (CLNPanel <> 'C' OR CLNPanel IS NULL OR CLNPanel = '')"
                    Call CalCLMAmt(strAppend1, strAppendCase, Record)
                    
                    'sql4 = " SELECT CLNType, Count(CASNumber) AS TotalClaimNP, SUM(CLMTotalApproved) AS TotalNonPanelAmt " & _
                          " From VIEW_CLMReport Where CLNType = '" & Trim(ClinicType) & "' AND (CLNPanel <> 'C' OR CLNPanel IS NULL OR CLNPanel = '')"
                  
                    'sql5 = " GROUP BY CLNType "
                    
                    'If Len(Trim(strAppend)) Then
                    '    sql4 = sql4 + strAppend + sql5
                    'Else
                    '    sql4 = sql4 + sql5
                    'End If
            
                    'rst3.Open sql4, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                                
                    'Record.SubItems(3) = IIf(IsNumeric(rst3!TotalClaimNP) = True, rst3!TotalClaimNP, 0)
                    'Record.SubItems(4) = Format(IIf(IsNumeric(rst3!TotalNonPanelAmt) = True, rst3!TotalNonPanelAmt, 0), "#,##0.00")
                    
                    'rst3.Close
                    'Set rst3 = Nothing
                
                End If
                
                Record.SubItems(5) = CDbl(Record.SubItems(1)) + CDbl(Record.SubItems(3))
                Record.SubItems(6) = Format(CDbl(Record.SubItems(2)) + CDbl(Record.SubItems(4)), "#,##0.00")
            
            End With
            lblCurrent = lblCurrent + 1
            rst1.MoveNext
        Loop
        
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CLNDetailListing(Optional strAppend As String)
Const METHOD_NAME = "CLNDetailListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT PAYCode, ClinicCode, CLNName, CASNumber, MBMNumber, CoverID, " & _
          " PatientName, CLMTotalActual, CLMTotalApproved, " & _
          " CLMRecoveryPaidbyM, DateVisit, " & _
          " CASStatus, Claimability, CASTreatReason1, MCDays From VIEW_CLMReport Where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "25"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
    
            Do Until rst1.EOF
            With rst1
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!ClinicCode) Then
                    Record.SubItems(1) = Trim(!ClinicCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(2) = Trim(!CLNName)
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!CasNumber) Then
                   Record.SubItems(3) = Trim(!CasNumber)
                End If
                                
                If Len(!MBMNumber) Then
                    Record.SubItems(4) = Trim(!MBMNumber)
                End If
                
                If Len(!coverid) Then
                    Record.SubItems(5) = !coverid
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(6) = Trim(!PatientName)
                End If
                
                If Len(!CLMTotalActual) Then
                    Record.SubItems(7) = Format(!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(8) = Format(!CLMTotalApproved, "#,##0.00")
                End If
                
                If Len(!CLMRecoveryPaidByM) Then
                    Record.SubItems(9) = Format(!CLMRecoveryPaidByM, "#,##0.00")
                End If
                
                If Len(!DateVisit) Then
                    Record.SubItems(10) = Format(!DateVisit, "dd/mm/yyyy")
                End If
                
                If Len(!CASStatus) Then
                   Record.SubItems(11) = !CASStatus
                End If
                
                If Len(!Claimability) Then
                    Record.SubItems(12) = !Claimability
                End If
                                
                If Len(!CASTreatReason1) Then
                    Record.SubItems(13) = Trim(!CASTreatReason1)
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(14) = Trim(!MCDays)
                End If
                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ResultListing(Optional strAppend As String)
Const METHOD_NAME = "ResultListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
        
    

    
    'sql = " SELECT PAYCode, CASNumber, MBMNumber, MBMPolicyNo, " & _
          " CoverID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, " & _
          " INSCode, PLNCode, ClinicCode, CLNName, CASTreatReason1, " & _
          " CLMTotalActual, CLMTotalApproved, MCDays, ReasonDesc, " & _
          " CLMPayablebyMedix2M, BordxRefNo, BordxDate, ClaimOpenDate, " & _
          " Claimability, CASStatus, MBMDataStatus, MBMStatus, DateVisit, " & _
          " MBMPayorEffDate, MBMPayorExpDate, MBMCLNPremium, " & _
          " RegisterDate, MBMIcBcPp, ICNo, CASDiagnosis5, MBMDepartment, INVPayTo, " & _
          " CLMNonGpFee, CLNType From VIEW_CLMReport Where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend '+ sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "93"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute

    Do
       Do Until rst1.EOF
                               
            With rst1
                    
                    If Len(!PAYCode) Then
                        Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(!CasNumber) Then
                       Record.SubItems(1) = Trim(!CasNumber)
                    End If
                    
                    If Len(!Bordxrefno) Then
                       Record.SubItems(2) = !Bordxrefno
                    End If
                    
                    If Len(!BordxDate) Then
                       Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMNumber) Then
                       Record.SubItems(4) = !MBMNumber
                    End If
                    
                    If Len(!MBMPolicyNo) Then
                       Record.SubItems(5) = Trim(!MBMPolicyNo)
                    End If
                                    
                    If Len(!coverid) Then
                       Record.SubItems(6) = !coverid
                    End If
                    
                    If Len(!MBMName) Then
                       Record.SubItems(7) = Trim(!MBMName)
                    End If
                    
                    If Len(!PatientName) Then
                       Record.SubItems(8) = Trim(!PatientName)
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                       Record.SubItems(9) = !MBMIcBcPp
                    End If
                    
                    If Len(!MBMSex) Then
                       Record.SubItems(10) = !MBMSex
                    End If
                    
                    If Len(!AgeCode) Then
                       Record.SubItems(11) = !AgeCode
                    End If
                    
                    If Len(!MBMDOB) Then
                        Record.SubItems(12) = Format(!MBMDOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!IcNo) Then
                       Record.SubItems(13) = !IcNo
                    End If
                    
                    If Len(!Sex) Then
                       Record.SubItems(14) = !Sex
                    End If
                    
                    If Len(!AgeBand) Then
                       Record.SubItems(15) = !AgeBand
                    End If
                    
                    If Len(!DOB) Then
                        Record.SubItems(16) = Format(!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INSCode) Then
                       Record.SubItems(17) = !INSCode
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(18) = !PLNCode
                    End If
                    
                    If Len(!ClinicCode) Then
                       Record.SubItems(19) = Trim(!ClinicCode)
                    End If
                    
                    If Len(!CLNName) Then
                        Record.SubItems(20) = Trim(!CLNName)
                    End If
                    
                    If !CASTreatReason1 = "016" Or !CASTreatReason1 = "032" Then
                        Record.SubItems(21) = !CASTreatReason1
                        Record.SubItems(22) = Trim(!CASDiagnosis5)
                    ElseIf Len(!CASTreatReason1) Then
                        Record.SubItems(21) = !CASTreatReason1
                        Record.SubItems(22) = IIf(Len(!ReasonDesc), !ReasonDesc, "")
                    Else
                        Record.SubItems(21) = "Empty"
                        Record.SubItems(22) = "Empty"
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(23) = Format(!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    Else
                       Record.SubItems(23) = "0.00"
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                        Record.SubItems(24) = Format(!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    Else
                        Record.SubItems(24) = "0.00"
                    End If
                                        
                    If Len(!CLMPayableByMedix2M) Then
                       
                       If !CLMPayableByMedix2M < 0 Then
                            Record.SubItems(25) = Format(!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                       Else
                            Record.SubItems(25) = "0.00"
                       End If
                    Else
                        Record.SubItems(25) = "0.00"
                    End If
                    
                    If Len(!MBMCLNPremium) Then
                        Record.SubItems(26) = Format(!MBMCLNPremium, "#,##0.00;(#,##0.00)")
                    Else
                        Record.SubItems(26) = "0.00"
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(27) = !CASStatus
                    End If
                    
                    If Len(!Claimability) Then
                       Record.SubItems(28) = !Claimability
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(29) = !MBMStatus
                    End If
                    
                    If Len(!MBMDataStatus) Then
                       Record.SubItems(30) = !MBMDataStatus
                    End If
                    
                    If Len(!DateVisit) Then
                        Record.SubItems(31) = Format(!DateVisit, "dd/mm/yyyy")
                    End If
                    
                    If Len(!RegisterDate) Then
                        Record.SubItems(32) = Format(!RegisterDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MCDays) Then
                        Record.SubItems(33) = !MCDays
                    End If
                    
                    If Len(!ClaimOpenDate) Then
                        Record.SubItems(34) = Format(!ClaimOpenDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(35) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(36) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMDepartment) Then
                        Record.SubItems(37) = !MBMDepartment
                    End If
                    
                    If Len(!INVPayTo) Then
                        Record.SubItems(38) = !INVPayTo
                    End If
                    
                    Record.SubItems(41) = IIf(Len(!ClmNonGPFee), Format(!ClmNonGPFee, "#,##0.00;(#,##0.00)"), "0.00")
                    
                    If rst1!CLNType = "S" Then
                        Record.SubItems(39) = "0.00"
                        Record.SubItems(40) = Format(!CLMTotalActual - IIf(IsNumeric(!ClmNonGPFee) = True, !ClmNonGPFee, 0), "##,#0.00")
                    Else
                        Record.SubItems(39) = Format(!CLMTotalActual - IIf(IsNumeric(!ClmNonGPFee) = True, !ClmNonGPFee, 0), "##,#0.00")
                        Record.SubItems(40) = "0.00"
                    End If
                    
                    If Len(!InvoiceNo) Then
                        Record.SubItems(42) = !InvoiceNo
                    End If
                    
                    If Len(!MBMGroupCompany) Then
                        Record.SubItems(43) = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
                        
                    End If
                    
                     If Len(!ReasonDesc11) Then
                        Record.SubItems(44) = IIf(Len(!ReasonDesc11), !ReasonDesc11, "")
                    End If
                    
                     If Len(!ReasonDesc22) Then
                        Record.SubItems(45) = IIf(Len(!ReasonDesc22), !ReasonDesc22, "")
                    End If
                    
                     If Len(!ReasonDesc33) Then
                        Record.SubItems(46) = IIf(Len(!ReasonDesc33), !ReasonDesc33, "")
                    End If
                    
                     If Len(!ReasonDesc44) Then
                        Record.SubItems(47) = IIf(Len(!ReasonDesc44), !ReasonDesc44, "")
                    End If
                    
                     If Len(!ReasonDesc55) Then
                        Record.SubItems(48) = IIf(Len(!ReasonDesc55), !ReasonDesc55, "")
                    End If
                    
                     If Len(!RegisterBy) Then
                        Record.SubItems(49) = IIf(Len(!RegisterBy), !RegisterBy, "")
                    End If
                    
                     If Len(!LastUpdateBy) Then
                        Record.SubItems(50) = IIf(Len(!LastUpdateBy), !LastUpdateBy, "")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ResultListing2(Optional strAppend As String)
Const METHOD_NAME = "ResultListing2"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
        
    ListView2.ListItems.Clear
    Current = 0
    total = 0
    
    'sql = " SELECT PAYCode, CASNumber, MBMNumber, MBMPolicyNo, " & _
          " CoverID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, " & _
          " INSCode, PLNCode, ClinicCode, CLNName, CASTreatReason1, ReasonDesc, " & _
          " CLMTotalActual, CLMTotalApproved, Remarks, MCDays, " & _
          " CLMPayablebyMedix2M, BordxRefNo, BordxDate, ClaimOpenDate, " & _
          " Claimability, CASStatus, MBMDataStatus, MBMStatus, DateVisit, " & _
          " MBMPayorEffDate, MBMPayorExpDate, MBMCLNPremium, " & _
          " RegisterDate, MBMIcBcPp, ICNo, CASDiagnosis5 From VIEW_CLMReport Where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend '+ sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "94"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
              
       Do Until rst1.EOF
       
       'total = rst1.RecordCount
       'lblTotalRecord = total
        
                        
            With rst1
                    
                    If Len(!PAYCode) Then
                        Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(!CasNumber) Then
                       Record.SubItems(1) = Trim(!CasNumber)
                    End If
                    
                    If Len(!Bordxrefno) Then
                       Record.SubItems(2) = !Bordxrefno
                    End If
                    
                    If Len(!BordxDate) Then
                       Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMNumber) Then
                       Record.SubItems(4) = !MBMNumber
                    End If
                    
                    If Len(!MBMPolicyNo) Then
                       Record.SubItems(5) = Trim(!MBMPolicyNo)
                    End If
                                    
                    If Len(!coverid) Then
                       Record.SubItems(6) = !coverid
                    End If
                    
                    If Len(!MBMName) Then
                       Record.SubItems(7) = Trim(!MBMName)
                    End If
                    
                    If Len(!PatientName) Then
                       Record.SubItems(8) = Trim(!PatientName)
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                       Record.SubItems(9) = !MBMIcBcPp
                    End If
                    
                    If Len(!MBMSex) Then
                       Record.SubItems(10) = !MBMSex
                    End If
                    
                    If Len(!AgeCode) Then
                       Record.SubItems(11) = !AgeCode
                    End If
                    
                    If Len(!MBMDOB) Then
                        Record.SubItems(12) = Format(!MBMDOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!IcNo) Then
                       Record.SubItems(13) = !IcNo
                    End If
                    
                    If Len(!Sex) Then
                       Record.SubItems(14) = !Sex
                    End If
                    
                    If Len(!AgeBand) Then
                       Record.SubItems(15) = !AgeBand
                    End If
                    
                    If Len(!DOB) Then
                        Record.SubItems(16) = Format(!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INSCode) Then
                       Record.SubItems(17) = !INSCode
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(18) = !PLNCode
                    End If
                    
                    If Len(!ClinicCode) Then
                       Record.SubItems(19) = Trim(!ClinicCode)
                    End If
                    
                    If Len(!CLNName) Then
                       Record.SubItems(20) = Trim(!CLNName)
                    End If
                    
                    If !CASTreatReason1 = "016" Or !CASTreatReason1 = "032" Then
                        Record.SubItems(21) = !CASTreatReason1
                        Record.SubItems(22) = Trim(!CASDiagnosis5)
                    ElseIf Len(!CASTreatReason1) Then
                        Record.SubItems(21) = !CASTreatReason1
                        Record.SubItems(22) = Trim(!ReasonDesc)
                    Else
                        Record.SubItems(21) = "Empty"
                        Record.SubItems(22) = "Empty"
                    End If
                    
                    'If Len(!CASTreatReason1) Then
                       'Record.SubItems(21) = !CASTreatReason1
                    'Else
                        'Record.SubItems(21) = "Empty"
                    'End If
                    
                    'If Len(!ReasonDesc) Then
                       'Record.SubItems(22) = Trim(!ReasonDesc)
                    'Else
                        'Record.SubItems(22) = "Empty"
                    'End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(23) = Format(!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    Else
                       Record.SubItems(23) = "0.00"
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                       Record.SubItems(24) = Format(!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    Else
                       Record.SubItems(24) = "0.00"
                    End If
                                        
                    If Len(!CLMPayableByMedix2M) Then
                       
                       If !CLMPayableByMedix2M < 0 Then
                            Record.SubItems(25) = Format(!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                       Else
                            Record.SubItems(25) = "0.00"
                       End If
                    Else
                        Record.SubItems(25) = "0.00"
                    End If
                    
                    If Len(!MBMCLNPremium) Then
                       Record.SubItems(26) = Format(!MBMCLNPremium, "#,##0.00;(#,##0.00)")
                    Else
                       Record.SubItems(26) = "0.00"
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(27) = !CASStatus
                    End If
                    
                    If Len(!Claimability) Then
                       Record.SubItems(28) = !Claimability
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(29) = !MBMStatus
                    End If
                    
                    If Len(!MBMDataStatus) Then
                       Record.SubItems(30) = !MBMDataStatus
                    End If
                    
                    If Len(!DateVisit) Then
                        Record.SubItems(31) = Format(!DateVisit, "dd/mm/yyyy")
                    End If
                    
                    If Len(!RegisterDate) Then
                        Record.SubItems(32) = Format(!RegisterDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MCDays) Then
                        Record.SubItems(33) = !MCDays
                    End If
                    
                    If Len(!ClaimOpenDate) Then
                        Record.SubItems(34) = Format(!ClaimOpenDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(35) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(36) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!Remarks) Then
                        Record.SubItems(37) = Trim(!Remarks)
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ResultListing3(Optional strAppend As String)
Const METHOD_NAME = "ResultListing3"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
        
    ListView2.ListItems.Clear
    Current = 0
    total = 0
    
    'sql = " SELECT Count(CASNumber) AS NoofVisits, MBMNumber, " & _
          " PatientName, PLNCode, SUM(CLMTotalApproved) AS AmtClaims, " & _
          " ICNo From VIEW_CLMReport Where 0 = 0 "
    
    'sql1 = " Group by MBMNumber, ICNo, PatientName, PLNCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql1
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "95"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
       Do Until rst1.EOF
       
            With rst1
                    
                    If Len(!NoofVisits) Then
                        Set Record = ListView2.ListItems.Add(, , !NoofVisits)
                    End If
                    
                    If Len(!IcNo) Then
                       Record.SubItems(1) = Trim(!IcNo)
                    End If
                    
                    If Len(!MBMNumber) Then
                       Record.SubItems(2) = Trim(!MBMNumber)
                    End If
                    
                    If Len(!PatientName) Then
                       Record.SubItems(3) = Trim(!PatientName)
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(4) = Trim(!PLNCode)
                    End If
                    
                    If Len(!AmtClaims) Then
                       Record.SubItems(5) = Format(!AmtClaims, "#,##0.00;(#,##0.00)")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub INSListing(Optional strAppend As String)
Const METHOD_NAME = "INSListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, INSCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved  " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql1 = " Group By PAYCode, INSCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql1
    'Else
    '    sql = sql + sql1
    'End If
    
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "67"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
        Do Until rst1.EOF
        
            With rst1
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!INSCode) Then
                    Record.SubItems(1) = !INSCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst1.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub StateListing(Optional strAppend As String)
Const METHOD_NAME = "StateListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
       
    'sql = " SELECT PAYCode, STACode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By PAYCode, STACode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
            
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "98"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!STACode) Then
                    Record.SubItems(1) = Trim(!STACode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub PayorListing(Optional strAppend As String)
Const METHOD_NAME = "PayorListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
      
    'sql = " SELECT PAYCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By PAYCode"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
     
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "80"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(2) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(3) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub ListViewHeader5()
     Const METHOD_NAME = "ListViewHeader5"
    On Error GoTo debugs
If CboSelection.Text = "By Clinic" Then
    ListView2.ColumnHeaders(1).Text = "Treatment Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Clinic Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Clinic Name"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "No of Visit"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "MC Days"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf CboSelection.Text = "Top 20 by Claims" Then
    ListView2.ColumnHeaders(1).Text = "Treatment Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "Top 20 by Amount" Then
    ListView2.ColumnHeaders(1).Text = "Treatment Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf CboSelection.Text = "General Summary" Then
    ListView2.ColumnHeaders(1).Text = "Treatment Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Visit"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "MC Days"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'********************** List Ageband Results *******************************
Private Sub AgeGroupListing(Optional strAppend As String)
Const METHOD_NAME = "AgeGroupListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY Ageband "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "14"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
        Do Until rst2.EOF

                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!AgeBand)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(2) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(3) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    'End If
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Ageband Results *******************************

'********************** List Ageband By Gender Results *******************************
Private Sub AgeGenderListing(Optional strAppend As String)
Const METHOD_NAME = "AgeGenderListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim strCount
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT Sex, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY Sex, Ageband "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "13"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
            
                If Len(rst2!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!Sex)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = rst2!AgeBand
                   'AgeCode = rst2!AgeBand
                Else
                   Record.SubItems(1) = "Empty"
                   'AgeCode = ""
                End If
                                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(3) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Ageband By Gender Results *******************************

'********************** List Ageband By Insured Type Results *******************************
Private Sub AgeInsuredListing(Optional strAppend As String)
Const METHOD_NAME = "AgeInsuredListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
       
    'sql = " SELECT PAYCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY PAYCode, INSCode, Ageband "
        
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "15"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!INSCode) Then
                   Record.SubItems(1) = !INSCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                              
                If Len(!AgeBand) Then
                   Record.SubItems(2) = !AgeBand
                Else
                    Record.SubItems(2) = "Empty"
                End If
                         
                If Len(!TotalCases) Then
                   Record.SubItems(3) = !TotalCases
                End If
                                
                If Len(!TotalActualAmt) Then
                   Record.SubItems(4) = Format(!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(!TotalApprovedAmt, "#,##0.00")
                End If
                
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Ageband By Insured Type Results *******************************

'********************** List Ageband By Race Type Results *******************************
Private Sub AgeRaceListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY RACCode, Ageband "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "21"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
                If Len(rst2!RACCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!RACCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = rst2!AgeBand
                Else
                   Record.SubItems(1) = "Empty"
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(3) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub
'********************** List Ageband By Race Type Results *******************************

'********************** List Ageband By Plan Type Results *******************************
Private Sub AgePlanListing(Optional strAppend As String)
Const METHOD_NAME = "AgePlanListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT INSCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY INSCode, PLNCode, Ageband "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "20"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!PLNCode) Then
                   Record.SubItems(1) = rst2!PLNCode
                End If
                    
                If Len(rst2!AgeBand) Then
                   Record.SubItems(2) = rst2!AgeBand
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Ageband By Plan Type Results *******************************


'********************** List Gender Results *******************************
Private Sub SexListing(Optional strAppend As String)
Const METHOD_NAME = "SexListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
    
    'sql = " SELECT SEX, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By SEX "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
      
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "100"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
            With rst2
                If Len(!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , !Sex)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                               
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(2) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(3) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Gender Results *******************************

'********************** List Gender By Ageband Results *******************************
Private Sub SexAgebandListing(Optional strAppend As String)
Const METHOD_NAME = "SexAgebandListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT SEX, Ageband, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By SEX, Ageband "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
      
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "99"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , !Sex)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!AgeBand) Then
                    Record.SubItems(1) = !AgeBand
                Else
                    Record.SubItems(1) = "Empty"
                End If
                    
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Gender By Ageband Results *******************************

'********************** List Gender By Insured Type Results *******************************
Private Sub SexInsuredListing(Optional strAppend As String)
Const METHOD_NAME = "SexInsuredListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
            
   'sql = " SELECT PAYCode, INSCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " Group By PAYCode, INSCode, Sex "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
      
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "69"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!INSCode) Then
                    Record.SubItems(1) = !INSCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Sex) Then
                    Record.SubItems(2) = !Sex
                Else
                    Record.SubItems(2) = "Empty"
                End If

                If Len(!TotalCases) Then
                    Record.SubItems(3) = !TotalCases
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                    
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Gender By Insured Type Results *******************************

'********************** List Gender By Race Results *******************************
Private Sub SexRaceListing(Optional strAppend As String)
Const METHOD_NAME = "SexRaceListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT RACCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By RACCode, Sex "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
      
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "102"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!RACCode) Then
                    Set Record = ListView2.ListItems.Add(, , !RACCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Sex) Then
                    Record.SubItems(1) = !Sex
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                    
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Gender By Race Results *******************************

'********************** List Gender By Plan Results *******************************
Private Sub SexPlanListing(Optional strAppend As String)
Const METHOD_NAME = "SexPlanListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT INSCode, PLNCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By INSCode, PLNCode, Sex "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
      
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "101"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(1) = !PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                    
                If Len(!Sex) Then
                    Record.SubItems(2) = !Sex
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(3) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Gender By Plan Results *******************************

'********************** List Race Results *******************************
Private Sub RaceListing(Optional strAppend As String)
Const METHOD_NAME = "RaceListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
       
    'sql = " SELECT RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By RACCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
      
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "90"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    
        Do Until rst2.EOF
        
            With rst2
                If Len(!RACCode) Then
                    Set Record = ListView2.ListItems.Add(, , !RACCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                               
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(2) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(3) = Format(!TotalApproved, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Race Results *******************************

'********************** List Race By Ageband Results *******************************
Private Sub RaceAgebandListing(Optional strAppend As String)
Const METHOD_NAME = "RaceAgebandListing"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By RACCode, Ageband "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
      
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "89"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!RACCode) Then
                    Set Record = ListView2.ListItems.Add(, , !RACCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!AgeBand) Then
                    Record.SubItems(1) = !AgeBand
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Race By Ageband Results *******************************

'********************** List Race By Insured Type Results *******************************
Private Sub RaceInsuredListing(Optional strAppend As String)
Const METHOD_NAME = "RaceInsuredListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
            
    'sql = " SELECT PAYCode, INSCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By PAYCode, INSCode, RACCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "70"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!INSCode) Then
                    Record.SubItems(1) = !INSCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!RACCode) Then
                    Record.SubItems(2) = !RACCode
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(3) = !TotalCases
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Race By Insured Type Results *******************************

'********************** List Race By Insured Type Results *******************************
Private Sub InvoiceListing(Optional strAppend As String)
Const METHOD_NAME = "InvoiceListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
            
    strAppendCase = "109"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 900
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
            
                 'SELECT Count(ClinicCode) As Total, InvoiceNo, ClinicCode, CLNName,
                 'SUM(CLMTotalActual) as CLMTotalActual,CLMProcMth, CLMProcYR,OPBatchNo,CLNType
                If Len(!ClinicCode) Then
                    Set Record = ListView2.ListItems.Add(, , !ClinicCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(1) = !CLNName
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!total) Then
                    Record.SubItems(2) = !total
                Else
                    Record.SubItems(2) = "0"
                End If
                
                If Len(!InvoiceNo) Then
                    Record.SubItems(3) = !InvoiceNo
                Else
                    Record.SubItems(3) = ""
                End If
                    
                If Len(!InvoiceDate) Then
                    Record.SubItems(4) = !InvoiceDate
                Else
                    Record.SubItems(4) = ""
                End If
                If Len(!CLMProcMth) Then
                    Record.SubItems(5) = !CLMProcMth & "/" & !CLMProcYR
                End If
                
                If Len(!CLMTotalActual) Then
                    Record.SubItems(6) = Format(!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(!OPBatchNo) Then
                    Record.SubItems(8) = !OPBatchNo
                Else
                    Record.SubItems(8) = ""
                End If
                
                If Trim(!ClinicCode) = "NP" Then
                    Record.SubItems(9) = "NP"
                Else
                    Record.SubItems(9) = !CLNType
                End If

            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Race By Insured Type Results *******************************

'********************** List Race By Gender Results *******************************
Private Sub RaceSexListing(Optional strAppend As String)
Const METHOD_NAME = "RaceSexListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT RACCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By RACCode, Sex "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "92"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    
        Do Until rst2.EOF
        
            With rst2
                If Len(!RACCode) Then
                    Set Record = ListView2.ListItems.Add(, , !RACCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Sex) Then
                    Record.SubItems(1) = !Sex
                Else
                    Record.SubItems(1) = "Empty"
                End If
                              
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Race By Sex Results *******************************

'********************** List Race By Plan Results *******************************
Private Sub RacePlanListing(Optional strAppend As String)
Const METHOD_NAME = "RacePlanListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT INSCode, PLNCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By INSCode, PLNCode, RACCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "91"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    
        Do Until rst2.EOF
        
            With rst2
                If Len(!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , !INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(1) = !PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                    
                If Len(!RACCode) Then
                    Record.SubItems(2) = !RACCode
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(3) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Race By Plan Results *******************************

'********************** List Agent Detail Results *******************************
Private Sub AgentListing(Optional strAppend As String)
Const METHOD_NAME = "AgentListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim StrSql As String
Dim FirstICDCode As String
Dim FirstDescription As String
        
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT PAYCode, MBMAgentCode, CASNumber, MBMNumber, " & _
          " CoverID, PatientName, CLMTotalActual, " & _
          " DateVisit, CLMTotalApproved, CLMRecoveryPaidbyM, " & _
          " CASTreatReason1, InsCode, PLNCode, CASStatus, MCDays " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If
            
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "17"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    Do
              
       Do Until rst2.EOF
                      
        FirstICDCode = ""
        FirstDescription = ""
                                 
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!MBMAgentCode) Then
                    Record.SubItems(1) = Trim(rst2!MBMAgentCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(rst2!CasNumber) Then
                   Record.SubItems(2) = Trim(rst2!CasNumber)
                End If
                
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(3) = rst2!MBMNumber
                End If
                                
                If Len(rst2!coverid) Then
                   Record.SubItems(4) = rst2!coverid
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(5) = Trim(rst2!PatientName)
                End If
                                
                If Len(rst2!CLMTotalActual) Then
                   Record.SubItems(7) = Format(rst2!CLMTotalActual, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                   Record.SubItems(8) = Format(rst2!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                Else
                   Record.SubItems(8) = "0.00"
                End If
                
                If Len(rst2!CLMRecoveryPaidByM) Then
                   Record.SubItems(9) = Format(rst2!CLMRecoveryPaidByM, "#,##0.00;(#,##0.00)")
                Else
                   Record.SubItems(9) = "0.00"
                End If
                
                If Len(rst2!CASTreatReason1) Then
                   Record.SubItems(10) = rst2!CASTreatReason1
                   FirstICDCode = Record.SubItems(10)
                   'sql3 = " SELECT ReasonCode, ReasonDesc From TreatmentReason Where ReasonCode = '" & Trim(FirstICDCode) & "'"
                   'rst3.Open sql3, MediConnCISDB, adOpenKeyset, adLockOptimistic
                   StrSql = " EXEC CISClaimTracking '" & Trim(FirstICDCode) & "', 10"
                   rst3.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                   
                    Do While Not rst3.EOF
                        FirstDescription = Trim(rst3!ReasonDesc)
                        rst3.MoveNext
                    Loop
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(FirstDescription) Then
                   Record.SubItems(11) = Trim(FirstDescription)
                End If
                
                                                    
                If Len(rst2!DateVisit) Then
                    Record.SubItems(12) = Format(rst2!DateVisit, "dd/mm/yyyy")
                End If
                
                If Len(rst2!MCDays) Then
                   Record.SubItems(13) = rst2!MCDays
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(14) = rst2!INSCode
                End If
                
                If Len(rst2!PLNCode) Then
                   Record.SubItems(15) = rst2!PLNCode
                End If
                                    
                If Len(rst2!CASStatus) Then
                   Record.SubItems(16) = rst2!CASStatus
                End If
                
                Current = Current + 1
                lblCurrent = Current
            
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Agent Detail Results *******************************

'********************** List Agent Summary Results *******************************
Private Sub AgentSumListing(Optional strAppend As String)
Const METHOD_NAME = "AgentSumListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
    
    If CboSelection.Text = "General Summary" Then
        strAppendCase = "18"
        'sql = " SELECT PAYCode, MBMAgentCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
              " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
              " From VIEW_CLMReport Where 0 = 0 "
          
        'sql2 = " Group By PAYCode, MBMAgentCode "
    
    ElseIf CboSelection.Text = "By Top 50 Claims" Then
        strAppendCase = "19"
        'sql = " SELECT Top 50 PAYCode, MBMAgentCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
              " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
              " From VIEW_CLMReport Where 0 = 0 "
          
        'sql2 = " Group By PAYCode, MBMAgentCode ORDER BY TotalCases Desc "
    
    End If
       
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!MBMAgentCode) Then
                    Record.SubItems(1) = Trim(rst2!MBMAgentCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                                         
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
       
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Agent Summary Results *******************************

'********************** List Plan Summary Results *******************************
Private Sub PlanListing(Optional strAppend As String)
Const METHOD_NAME = "PlanListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
    
    'sql = " SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
        
    'sql2 = " Group By PAYCode, PLNCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "84"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(1) = !PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                               
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub PlanPremiumListing(Optional strAppend As String)
Const METHOD_NAME = "PlanPremiumListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
       
    'sql = " SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(MBMCLNBasicPrem) AS BasicPremium, " & _
          " SUM(MBMCLNPremium) AS Premium From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By PAYCode, PLNCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "85"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(1) = !PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                               
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!BasicPremium) Then
                    Record.SubItems(5) = Format(!BasicPremium, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!Premium) Then
                    Record.SubItems(6) = Format(!Premium, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Plan Summary Results *******************************

'********************** List Plan Top 10 Clinic Results *******************************
Private Sub PlanTopHosListing(Optional strAppend As String)
Const METHOD_NAME = "PlanTopHosListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT Top 10 PAYCode, PLNCode, ClinicCode, CLNName, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By PAYCode, PLNCode, ClinicCode, CLNName ORDER BY TotalCases Desc"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
            
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "87"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(1) = Trim(!PLNCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!ClinicCode) Then
                    Record.SubItems(2) = Trim(!ClinicCode)
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(3) = Trim(!CLNName)
                Else
                    Record.SubItems(3) = "Empty"
                End If
                           
                If Len(!TotalCases) Then
                    Record.SubItems(4) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(5) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(6) = Format(!TotalApproved, "#,##0.00")
                End If
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Plan Top 10 Clinic Results *******************************

'********************** List Plan Top 10 Final Diag Results *******************************
Private Sub PlanTopDiagListing(Optional strAppend As String)
Const METHOD_NAME = "PlanTopDiagListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim StrSql As String
    
    lblCurrent = 0
    DIagCode = ""
    ListView2.ListItems.Clear
        
    'sql = " SELECT Top 10 PAYCode, PLNCode, CASTreatReason1, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By PAYCode, PLNCode, CASTreatReason1 ORDER BY TotalCases Desc"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
            
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "86"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(1) = !PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(rst2!CASTreatReason1) Then
                    Record.SubItems(2) = rst2!CASTreatReason1
                    DIagCode = rst2!CASTreatReason1
                    
                    'sql3 = " SELECT ReasonCode, ReasonDesc From TreatmentReason Where ReasonCode = '" & Trim(DIagCode) & "'"
                    'rst3.Open sql3, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
                    StrSql = "EXEC CISClaimTracking '" & Trim(DIagCode) & "',10"
                    rst3.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                   
                   If rst3.EOF = True Then
                        Record.SubItems(3) = ""
                        rst3.Close
                        Set rst3 = Nothing
                   Else
                        Record.SubItems(3) = Trim(rst3!ReasonDesc)
                        rst3.Close
                        Set rst3 = Nothing
                   End If
                Else
                    Record.SubItems(2) = "Empty"
                    DIagCode = ""
                End If
                    
                If Len(!TotalCases) Then
                    Record.SubItems(4) = !TotalCases
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(5) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(6) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Plan Top 10 Final Diag Results *******************************

'********************** List Plan Type by Ageband Results *******************************
Private Sub PlanAgebandListing(Optional strAppend As String)
Const METHOD_NAME = "PlanAgebandListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
    
    'sql = " SELECT PAYCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY PAYCode, PLNCode, Ageband "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "81"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

        
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!PLNCode) Then
                   Record.SubItems(1) = !PLNCode
                Else
                   Record.SubItems(1) = "Empty"
                End If
                    
                If Len(!AgeBand) Then
                   Record.SubItems(2) = !AgeBand
                Else
                    Record.SubItems(2) = "Empty"
                End If
                                    
                If Len(!TotalCases) Then
                   Record.SubItems(3) = !TotalCases
                End If
                                
                If Len(!TotalActualAmt) Then
                   Record.SubItems(4) = Format(!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(!TotalApprovedAmt, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Plan Type By Ageband Results *******************************

'********************** List Plan Type by Insured Results *******************************
Private Sub PlanInsuredListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, PLNCode, INSCode, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY PAYCode, PLNCode, INSCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "68"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!INSCode) Then
                   Record.SubItems(1) = !INSCode
                Else
                   Record.SubItems(1) = "Empty"
                End If
                    
                If Len(!PLNCode) Then
                   Record.SubItems(2) = !PLNCode
                Else
                   Record.SubItems(2) = "Empty"
                End If
                    
                If Len(!TotalCases) Then
                   Record.SubItems(3) = !TotalCases
                End If
                                
                If Len(!TotalActualAmt) Then
                   Record.SubItems(4) = Format(!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub PlanInsuredListing2(Optional strAppend As String)
Const METHOD_NAME = "PlanInsuredListing2"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, PLNCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY PAYCode, PLNCode, INSCode, Ageband "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "83"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!INSCode) Then
                   Record.SubItems(1) = !INSCode
                Else
                    Record.SubItems(1) = "Empty"
                End If
                    
                If Len(!PLNCode) Then
                   Record.SubItems(2) = !PLNCode
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!AgeBand) Then
                   Record.SubItems(3) = !AgeBand
                Else
                    Record.SubItems(3) = "Empty"
                End If
                    
                If Len(!TotalCases) Then
                   Record.SubItems(4) = !TotalCases
                End If
                                
                If Len(!TotalActualAmt) Then
                   Record.SubItems(5) = Format(!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
'********************** List Plan Type By Insured Results *******************************

'********************** List Plan By Panel Clinic *******************************
Private Sub PlanByPanel(Optional strAppend As String)
Const METHOD_NAME = "PlanByPanel"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
    PLNType = ""
        
    'sql = " SELECT PLNCode, SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql1 = " GROUP BY PLNCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql1
    'Else
    '    sql = sql + sql1
    'End If
    
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "82"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
        Do Until rst1.EOF
        
            With rst1
                If Len(!PLNCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PLNCode)
                    PLNType = !PLNCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!TotalApprovedAmt) Then
                   Record.SubItems(1) = Format(!TotalApprovedAmt, "#,##0.00")
                End If
                
                strAppend1 = strAppend + " AND PLNCode = '" & Trim(PLNType) & "' AND CLNPanel = 'C'"
                strAppendCase = "12"
                Call CalPanelAmt(strAppend1, strAppendCase, Record)
                
                strAppend1 = strAppend + " AND PLNCode = '" & Trim(PLNType) & "' AND (CLNPanel <> 'C' OR CLNPanel IS NULL OR CLNPanel = '')"
                strAppendCase = "33"
                Call CalPanelAmt(strAppend1, strAppendCase, Record)
                
            'sql2 = " SELECT PLNCode, SUM(CLMTotalApproved) AS TotalPanelAmt " & _
                  " From VIEW_CLMReport Where PLNCode = '" & Trim(PLNType) & "' AND CLNPanel = 'C'"
          
            'sql3 = " GROUP BY PLNCode "
            
            'If Len(Trim(strAppend)) Then
            '    sql2 = sql2 + strAppend + sql3
            'Else
            '    sql2 = sql2 + sql13
            'End If
    
            'rst2.Open sql2, MediConnCISWS, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
                        
            'Record.SubItems(2) = Format(IIf(IsNumeric(rst2!TotalPanelAmt) = True, rst2!TotalPanelAmt, 0), "#,##0.00")
            
            'rst2.Close
            'Set rst2 = Nothing
            
            'sql4 = " SELECT PLNCode, SUM(CLMTotalApproved) AS TotalNonPanelAmt " & _
                  " From VIEW_CLMReport Where PLNCode = '" & Trim(PLNType) & "' AND (CLNPanel <> 'C' OR CLNPanel IS NULL OR CLNPanel = '')"
          
            'sql5 = " GROUP BY PLNCode "
            
            'If Len(Trim(strAppend)) Then
            '    sql4 = sql4 + strAppend + sql5
            'Else
            '    sql4 = sql4 + sql5
            'End If
    
            'rst3.Open sql4, MediConnCISWS, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
                
            'Record.SubItems(3) = Format(IIf(IsNumeric(rst3!TotalNonPanelAmt) = True, rst3!TotalNonPanelAmt, 0), "#,##0.00")
            
            'rst3.Close
            'Set rst3 = Nothing
                    
            End With
            lblCurrent = lblCurrent + 1
            rst1.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'********************** List Plan By Panel Clinic *******************************

'********************** List Clinic Group Summary Results *******************************
Private Sub ClinicGrpSumListing(Optional strAppend As String)
Const METHOD_NAME = "ClinicGrpSumListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
    
    'sql = " SELECT ClinicGRPName, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY ClinicGRPName ORDER BY ClinicGRPName "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "22"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
        
        Do Until rst1.EOF
        
            With rst1
                If Len(!ClinicGrpName) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!ClinicGrpName))
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                    
                If Len(!TotalCases) Then
                   Record.SubItems(1) = !TotalCases
                End If
                                    
                If Len(!TotalActualAmt) Then
                   Record.SubItems(2) = Format(!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(!TotalApprovedAmt) Then
                   Record.SubItems(3) = Format(!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst1.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'********************** List Clinic Group Summary Results *******************************

Private Sub ICSummary(Optional strAppend As String)
Const METHOD_NAME = "ICSummary"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim strCount
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT ICNo, Count(CASNumber) AS TotalClaim, PatientName, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, SUM(CONVERT(Money,CLMPayableByMedix2M)) AS OSAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY ICNo, PatientName "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "66"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
               
        Do Until rst2.EOF
        
            With rst2
                If Len(!IcNo) Then
                    Set Record = ListView2.ListItems.Add(, , !IcNo)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(1) = Trim(!PatientName)
                End If
                
                If Len(!TotalClaim) Then
                   Record.SubItems(2) = Trim(!TotalClaim)
                End If
                
                If Len(!ActualAmt) Then
                   Record.SubItems(3) = Format(!ActualAmt, "#,##0.00")
                End If
                                        
                If Len(!ApprovedAmt) Then
                   Record.SubItems(4) = Format(!ApprovedAmt, "#,##0.00")
                End If
                
                If Len(!OSAmt) Then
                   Record.SubItems(5) = Format(!OSAmt, "#,##0.00")
                Else
                   Record.SubItems(5) = "0.00"
                End If
                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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

'********************** List Bordx No *******************************
Private Sub BordxNoListing(Optional strAppend As String)
Const METHOD_NAME = "BordxNoListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
        
    lblCurrent = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT Distinct BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " BordxClaimAmt From VIEW_CLMReport Where BordxRefNo IS NOT NULL "
          
    'sql2 = " Group By BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "7"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!Bordxrefno) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!Bordxrefno)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                    
                If Len(rst2!BordxDate) Then
                    Record.SubItems(1) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                               
                If Len(rst2!BordxNoClaims) Then
                    Record.SubItems(2) = Trim(rst2!BordxNoClaims)
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(5) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub BordxNoGrpListing(Optional strAppend As String)
Const METHOD_NAME = "BordxNoGrpListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
        
    lblCurrent = 0
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT Distinct MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " BordxClaimAmt From VIEW_CLMReport Where BordxRefNo IS NOT NULL "
          
    'sql2 = " Group By MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "8"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Bordxrefno) Then
                    Record.SubItems(1) = !Bordxrefno
                Else
                    Record.SubItems(1) = "Empty"
                End If
                    
                If Len(!BordxDate) Then
                    Record.SubItems(2) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                               
                If Len(!BordxNoClaims) Then
                    Record.SubItems(3) = Trim(!BordxNoClaims)
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                
                If Len(!BordxClaimAmt) Then
                    Record.SubItems(6) = Format(!BordxClaimAmt, "#,##0.00")
                End If
                
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub BordxNoDeptListing(Optional strAppend As String)
Const METHOD_NAME = "BordxNoDeptListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
Dim BordxNo As String
Dim ChkFlag As Boolean
        
    lblCurrent = 0
    BordxNo = ""
    ChkFlag = False
    
    ListView2.ListItems.Clear
            
    'sql = " SELECT Distinct MBMDepartment, MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " BordxClaimAmt From VIEW_CLMReport Where BordxRefNo IS NOT NULL "
          
    'sql2 = " Group By MBMDepartment, MBMGroupCompany, BordxRefNo, BordxDate, " & _
            " BordxNoClaims, BordxClaimAmt "
    
    'sql3 = " Order By BordxRefNo "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2 + sql3
    'Else
    '    sql = sql + sql2 + sql3
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "9"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                ChkFlag = IIf(StrComp(!Bordxrefno, BordxNo) = 0, True, False)
                
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMGroupCompany) Then
                    Record.SubItems(1) = !MBMGroupCompany
                Else
                    Record.SubItems(1) = "Empty"
                End If
                    
                If Len(!Bordxrefno) Then
                    Record.SubItems(2) = !Bordxrefno
                    BordxNo = !Bordxrefno
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!BordxDate) Then
                    Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                    
                If Len(!TotalActual) Then
                    Record.SubItems(5) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(6) = Format(!TotalApproved, "#,##0.00")
                End If
                    
                If ChkFlag = False Then
                    
                    If Len(!BordxNoClaims) Then
                        Record.SubItems(4) = Trim(!BordxNoClaims)
                    End If
                    
                    If Len(!BordxClaimAmt) Then
                        Record.SubItems(7) = Format(!BordxClaimAmt, "#,##0.00")
                    End If
                Else
                    Record.SubItems(4) = 0
                    Record.SubItems(7) = Format(0, "#,##0.00")
                End If
                
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub MemberSummaryByPayor(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, GRPCoCode, MBMNumber, PatientName, PLNCode, Count(CASNumber) AS TotalClaim, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY PAYCode, GRPCoCode, MBMNumber, PatientName, PLNCode "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "77"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

               
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!GRPCoCode) Then
                    Record.SubItems(1) = Trim(!GRPCoCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(2) = Trim(!MBMNumber)
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(3) = Trim(!PatientName)
                End If
                
                If Len(!PLNCode) Then
                    Record.SubItems(4) = Trim(!PLNCode)
                End If
                
                If Len(!TotalClaim) Then
                    Record.SubItems(5) = Trim(!TotalClaim)
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(6) = Trim(!MCDays)
                End If
                
                If Len(!ActualAmt) Then
                   Record.SubItems(7) = Format(!ActualAmt, "#,##0.00")
                End If
                                        
                If Len(!ApprovedAmt) Then
                   Record.SubItems(8) = Format(!ApprovedAmt, "#,##0.00")
                End If
                                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub MemberSummaryByGrpCompany(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMGroupCompany, MBMNumber, Count(CASNumber) AS TotalClaim, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY MBMGroupCompany, MBMNumber "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "76"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
               
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(1) = Trim(!MBMNumber)
                End If
                
                If Len(!TotalClaim) Then
                    Record.SubItems(2) = Trim(!TotalClaim)
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(3) = Trim(!MCDays)
                End If
                
                If Len(!ActualAmt) Then
                   Record.SubItems(4) = Format(!ActualAmt, "#,##0.00")
                End If
                                        
                If Len(!ApprovedAmt) Then
                   Record.SubItems(5) = Format(!ApprovedAmt, "#,##0.00")
                End If
                                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub MemberSummaryOS(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
            
        
    'sql = " SELECT MBMDepartment, MBMNumber, PatientName, CoverID, SUM(CLMTotalActual) AS ApprovedAmt, " & _
          " Count(CASNumber) AS TotalClaim, SUM(MCStatus) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY MBMDepartment, MBMNumber, PatientName, CoverID"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "78"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

               
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(1) = !MBMNumber
                End If
                
                If Len(!coverid) Then
                    Record.SubItems(2) = Trim(!coverid)
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(3) = Trim(!PatientName)
                End If
                
                If Len(!TotalClaim) Then
                    Record.SubItems(4) = Trim(!TotalClaim)
                Else
                    Record.SubItems(4) = "0"
                End If
                
                If Len(!MCTimes) Then
                    Record.SubItems(5) = Trim(!MCTimes)
                Else
                    Record.SubItems(5) = "0"
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(6) = Trim(!MCDays)
                Else
                    Record.SubItems(6) = "0"
                End If
                                        
                If Len(!ApprovedAmt) Then
                    Record.SubItems(7) = Format(!ApprovedAmt, "#,##0.00")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ReportMBMSummaryOS(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim MBMNo As String
Dim TotalClaim As Double
Dim MCTimes As Double
Dim MCDays As Double
Dim ApprovedAmt As Double
Dim BalLmt As Double
Dim Count As Double
Dim tmpTopupAmt As Double
    lblCurrent = 0
    tmpTopupAmt = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMDepartment, MBMNumber, MBMName, CoverID, AnnLimit, AnnLmtIND, SUM(CLMTotalActual) AS ApprovedAmt, " & _
          " Count(CASNumber) AS TotalClaim, SUM(MCStatus) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, MBMCLNBalLmt From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY MBMDepartment, MBMNumber, MBMName, CoverID, AnnLimit, AnnLmtIND, MBMCLNBalLmt"
    
    'sql3 = " ORDER BY MBMNumber"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2 + sql3
    'Else
    '    sql = sql + sql2 + sql3
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "5"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
            If StrComp(!MBMNumber, MBMNo) <> 0 Then
                BalLmt = 0
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(1) = !MBMNumber
                    MBMNo = !MBMNumber
                End If
                
                If Len(!MBMName) Then
                    Record.SubItems(2) = Trim(!MBMName)
                End If
                
                If Trim(!AnnLmtIND) = "Y" Then
                    If Len(!AnnLimit) Then
                        Record.SubItems(6) = Format(!AnnLimit, "#,##0.00")
                    Else
                        Record.SubItems(6) = "0.00"
                    End If
                Else
                    Record.SubItems(6) = "FULL"
                End If
                
                If Trim(!AnnLmtIND) = "Y" Then
                    If IsNumeric(!TopUpAmt) Then
                        Record.SubItems(7) = Format(!TopUpAmt, "#,##0.00")
                        tmpTopupAmt = Format(!TopUpAmt, "#,##0.00")
                    Else
                        Record.SubItems(7) = "0.00"
                        tmpTopupAmt = 0
                    End If
                Else
                    Record.SubItems(7) = "0.00"
                    tmpTopupAmt = 0
                End If
                
                If IsNumeric(!CLMMCOProcFee) Then
                    Record.SubItems(9) = Format(!CLMMCOProcFee, "#,##0.00")
                Else
                    Record.SubItems(9) = "0.00"
                End If
                
                
                If Trim(!AnnLmtIND) = "Y" Then
                    Record.SubItems(10) = Format(!MBMCLNBalLmt, "#,##0.00")
                    BalLmt = !MBMCLNBalLmt
                Else
                    Record.SubItems(10) = "FULL"
                    Record.SubItems(11) = "FULL"
                End If
                
                TotalClaim = 0
                MCTimes = 0
                MCDays = 0
                ApprovedAmt = 0
                lblCurrent = lblCurrent + 1
            End If
                
                If Len(!TotalClaim) Then
                    TotalClaim = TotalClaim + Trim(!TotalClaim)
                    Record.SubItems(3) = TotalClaim
                Else
                    Record.SubItems(3) = "0"
                End If
                
                If Len(!MCTimes) Then
                    MCTimes = MCTimes + Trim(!MCTimes)
                    Record.SubItems(4) = MCTimes
                Else
                    Record.SubItems(4) = "0"
                End If
                
                If Len(!MCDays) Then
                    MCDays = MCDays + Trim(!MCDays)
                    Record.SubItems(5) = MCDays
                Else
                    Record.SubItems(5) = "0"
                End If
                                        
                If Len(!ApprovedAmt) Then
                    ApprovedAmt = ApprovedAmt + Format(!ApprovedAmt, "#,##0.00")
                    Record.SubItems(8) = Format(ApprovedAmt, "#,##0.00")
                Else
                    Record.SubItems(8) = "0.00"
                End If
                                    
                'If Record.SubItems(6) <> 0 Then
                    If Record.SubItems(6) <> "FULL" Then
                        'count = ApprovedAmt + BalLmt
                        'If tmpTopupAmt < 0 Then
                        '    tmpTopupAmt = tmpTopupAmt * -1
                        'End If
                        
                        If ApprovedAmt > 0 Then
                            If tmpTopupAmt < 0 Then
                                If !AnnLimit > 0 Then
                                    Count = ((tmpTopupAmt * -1) + ApprovedAmt) / !AnnLimit
                                Else
                                    Count = 0
                                End If
                            Else
                                If !AnnLimit > 0 Then
                                    Count = ApprovedAmt / !AnnLimit
                                Else
                                    Count = 0
                                End If
                            End If
                            
                            Record.SubItems(11) = Format(Round(Count * 100, 2), "#,##0.00")
                        Else
                            Record.SubItems(11) = Format(0, "#,##0.00")
                        End If
                    Else
                        Record.SubItems(11) = "FULL"
                    End If
                'Else
                '    Record.SubItems(10) = Format(0, "#,##0.00")
                'End If
                    
            End With
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ReportMBMSummaryOS1(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim MBMNo As String
Dim TotalClaim As Double
Dim MCTimes As Double
Dim MCDays As Double
Dim ApprovedAmt As Double
Dim BalLmt As Double
Dim Count As Double
Dim tmpTopupAmt As Double
Dim tmpAnnLmt As Double

    tmpTopupAmt = 0
    lblCurrent = 0
    tmpAnnLmt = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMDepartment, MBMNumber, PatientName, CoverID, AnnLimit, AnnLmtIND, SUM(CLMTotalActual) AS ApprovedAmt, " & _
          " Count(CASNumber) AS TotalClaim, SUM(MCStatus) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, MBMCLNBalLmt From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY MBMDepartment, PatientName, MBMName, CoverID, AnnLimit, AnnLmtIND, MBMCLNBalLmt"
    
    'sql3 = " ORDER BY PatientName"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2 + sql3
    'Else
    '    sql = sql + sql2 + sql3
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "44"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
            'If StrComp(!MBMNumber, MBMNo) <> 0 Then
                BalLmt = 0
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(1) = !MBMNumber
                    MBMNo = !MBMNumber
                End If
                
                If Len(!coverid) Then
                    Record.SubItems(2) = Trim(!coverid)
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(3) = Trim(!PatientName)
                End If
                
                If Trim(!AnnLmtIND) = "Y" Then
                    If Len(!AnnLimit) Then
                        Record.SubItems(7) = Format(!AnnLimit, "#,##0.00")
                    Else
                        Record.SubItems(7) = "0.00"
                    End If
                Else
                    Record.SubItems(7) = "FULL"
                End If
                
                If Trim(!AnnLmtIND) = "Y" Then
                    
                    If Len(!TopUpAmt) Then
                        Record.SubItems(8) = Format(!TopUpAmt, "#,##0.00")
                        tmpTopupAmt = Format(!TopUpAmt, "#,##0.00")
                    Else
                        Record.SubItems(8) = "0.00"
                        tmpTopupAmt = 0
                    End If
                Else
                    Record.SubItems(8) = "FULL"
                    tmpTopupAmt = 0
                End If
                If Len(!CLMMCOProcFee) Then
                    ApprovedAmt = ApprovedAmt + Format(!CLMMCOProcFee, "#,##0.00")
                    Record.SubItems(10) = Format(!CLMMCOProcFee, "#,##0.00")
                Else
                    Record.SubItems(10) = "0.00"
                End If
                
                If Trim(!AnnLmtIND) = "Y" Then
                    Record.SubItems(11) = Format(!MBMCLNBalLmt, "#,##0.00")
                    BalLmt = !MBMCLNBalLmt
                Else
                    Record.SubItems(11) = "FULL"
                    Record.SubItems(12) = "FULL"
                End If
                
                TotalClaim = 0
                MCTimes = 0
                MCDays = 0
                ApprovedAmt = 0
                lblCurrent = lblCurrent + 1
            'End If
                
                If Len(!TotalClaim) Then
                    TotalClaim = TotalClaim + Trim(!TotalClaim)
                    Record.SubItems(4) = TotalClaim
                Else
                    Record.SubItems(4) = "0"
                End If
                
                If Len(!MCTimes) Then
                    MCTimes = MCTimes + Trim(!MCTimes)
                    Record.SubItems(5) = MCTimes
                Else
                    Record.SubItems(5) = "0"
                End If
                
                If Len(!MCDays) Then
                    MCDays = MCDays + Trim(!MCDays)
                    Record.SubItems(6) = MCDays
                Else
                    Record.SubItems(6) = "0"
                End If
                                        
                If Len(!ApprovedAmt) Then
                    ApprovedAmt = ApprovedAmt + Format(!ApprovedAmt, "#,##0.00")
                    Record.SubItems(9) = Format(ApprovedAmt, "#,##0.00")
                Else
                    Record.SubItems(9) = "0.00"
                End If
                                    
                'If Record.SubItems(7) <> 0 Then
                    If Record.SubItems(10) <> "FULL" Then
                        If ApprovedAmt > 0 Then
                            If Len(!AnnLimit) Then
                                If tmpTopupAmt < 0 Then
                                    If !AnnLimit > 0 Then
                                        Count = ((tmpTopupAmt * -1) + ApprovedAmt) / !AnnLimit
                                    Else
                                        Count = 0
                                    End If
                                Else
                                    If !AnnLimit > 0 Then
                                        Count = ApprovedAmt / !AnnLimit
                                    Else
                                        Count = 0
                                    End If
                                End If
                                
                                Record.SubItems(12) = Format(Round(Count * 100, 2), "#,##0.00")
                            Else
                                Record.SubItems(12) = "0"
                            End If
                        Else
                            Record.SubItems(12) = Format(0, "#,##0.00")
                        End If
                    Else
                        Record.SubItems(1) = "FULL"
                    End If
                'Else
                '    Record.SubItems(11) = Format(0, "#,##0.00")
                'End If
                    
            End With
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub MemberSummary(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMNumber, COUNT(CASNumber) AS TotalClaim, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY MBMNumber "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "75"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

               
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMNumber) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMNumber)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!TotalClaim) Then
                    Record.SubItems(1) = Trim(!TotalClaim)
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(2) = Trim(!MCDays)
                End If
                
                If Len(!ActualAmt) Then
                   Record.SubItems(3) = Format(!ActualAmt, "#,##0.00")
                End If
                                        
                If Len(!ApprovedAmt) Then
                   Record.SubItems(4) = Format(!ApprovedAmt, "#,##0.00")
                End If
                                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ClaimsSameDate(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    strAppendCase = "110"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

               
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMNumber) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMNumber)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!coverid) Then
                    Record.SubItems(1) = Trim(!coverid)
                End If
                If Len(!PatientName) Then
                    Record.SubItems(2) = Trim(!PatientName)
                End If
                If Len(!ClinicCode) Then
                    Record.SubItems(3) = Trim(!ClinicCode)
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(4) = Trim(!CLNName)
                End If
                
                If Len(!DateVisit) Then
                    Record.SubItems(5) = Trim(!DateVisit)
                End If
                
                If Len(!TotalClaim) Then
                    Record.SubItems(6) = Trim(!TotalClaim)
                End If
                
                If Len(!ActualAmt) Then
                   Record.SubItems(7) = Format(!ActualAmt, "#,##0.00")
                End If
                                        
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ClaimsSameAmt(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    strAppendCase = "111"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

               
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMNumber) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMNumber)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!coverid) Then
                    Record.SubItems(1) = Trim(!coverid)
                End If
                If Len(!PatientName) Then
                    Record.SubItems(2) = Trim(!PatientName)
                End If
                If Len(!ClinicCode) Then
                    Record.SubItems(3) = Trim(!ClinicCode)
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(4) = Trim(!CLNName)
                End If
                
                If Len(!DateVisit) Then
                    Record.SubItems(5) = Trim(!DateVisit)
                End If
                
                If Len(!TotalClaim) Then
                    Record.SubItems(6) = Trim(!TotalClaim)
                End If
                
                If Len(!CLMTotalActual) Then
                   Record.SubItems(7) = Format(!CLMTotalActual, "#,##0.00")
                End If
                                                                            
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub
Private Sub UtilizationByIllness(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim StrSql As String
Dim strAppend1 As String

    DIagCode = ""
    lblCurrent = 0
    total = ""
    
    ListView2.ListItems.Clear
       
    'sql = " SELECT CASTreatReason1, Sum(CLMTotalActual) As CLMTotalActual, " & _
          " SUM(CLMNonGpFee) AS CLMNonGpFee From VIEW_CLMReport Where 0 = 0 " & strAppend & _
          " GROUP BY CASTreatReason1 "
    
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "6"

    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
           
        Do Until rst1.EOF
    
            With rst1
                               
                If Len(!CASTreatReason1) Then
                    Set Record = ListView2.ListItems.Add(, , !CASTreatReason1)
                    DIagCode = !CASTreatReason1
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    DIagCode = ""
                End If
                
                'sql2 = " SELECT ReasonCode, ReasonDesc From TreatmentReason Where ReasonCode = '" & Trim(DIagCode) & "'"
                'rst2.Open sql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly 'adOpenKeyset, adLockOptimistic
                StrSql = "EXEC CISClaimTracking '" & Trim(DIagCode) & "', 10"
                rst2.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                       
                If rst2.EOF = True Then
                    rst2.Close
                    Set rst2 = Nothing
                    Record.SubItems(1) = ""
                Else
                    Description = Trim(rst2!ReasonDesc)
                    Record.SubItems(1) = Trim(Description)
                    rst2.Close
                    Set rst2 = Nothing
                End If
                                
                If Len(!ClmNonGPFee) Then
                    Record.SubItems(9) = Format(!ClmNonGPFee, "##,#0.00")
                Else
                    Record.SubItems(9) = "0.00"
                End If
                
                If Len(!CLMTotalActual) Then
                    Record.SubItems(10) = Format(!CLMTotalActual, "##,#0.00")
                Else
                    Record.SubItems(10) = "0.00"
                End If
                
                '===========================================================================
                If DIagCode = "" Then
                    strAppend1 = strAppend + " AND CLNType = 'G' AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL) "
                    strAppendCase = "11"
                    'sql5 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where CLNType = 'G' AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL) " & strAppend & _
                           " GROUP BY CASTreatReason1 "
                Else
                    strAppend1 = strAppend + " AND CLNType = 'G' AND CASTreatReason1 = '" & Trim(DIagCode) & "'"
                    strAppendCase = "11"
                    'sql5 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where CLNType = 'G' AND CASTreatReason1 = '" & Trim(DIagCode) & "'" & strAppend & _
                           " GROUP BY CASTreatReason1 "
                End If
                
                'rst5.Open sql5, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                Call CalGPAmt(strAppend1, strAppendCase, Record)
                       
                'If rst5.EOF = True Then
                '    rst5.Close
                '    Set rst5 = Nothing
                '    Record.SubItems(2) = "0"
                '    Record.SubItems(6) = "0.00"
                'Else
                '    Record.SubItems(2) = rst5!Total
                '    Record.SubItems(6) = Format(rst5!TotalApproved, "##,#0.00")
                '    rst5.Close
                '    Set rst5 = Nothing
                'End If
                '===========================================================================
                
                If Record.SubItems(6) > "0.00" Or Record.SubItems(2) > "0" Then
                    Record.SubItems(12) = Round((Record.SubItems(6) / Record.SubItems(2)), 2)
                Else
                    Record.SubItems(12) = "0"
                End If
                
                If DIagCode = "" Then
                    strAppend1 = strAppend + " AND CLNType = 'S' AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL) "
                    strAppendCase = "11"
                    'sql3 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where CLNType = 'S' AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL) " & strAppend & _
                           " GROUP BY CASTreatReason1 "
                Else
                    strAppend1 = strAppend + " AND CLNType = 'S' AND CASTreatReason1 = '" & Trim(DIagCode) & "'"
                    strAppendCase = "11"
                    'sql3 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where CLNType = 'S' AND CASTreatReason1 = '" & Trim(DIagCode) & "'" & strAppend & _
                           " GROUP BY CASTreatReason1 "
                End If
          
                'rst3.Open sql3, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                Call CalSPAmt(strAppend1, strAppendCase, Record)
                       
                'If rst3.EOF = True Then
                '    rst3.Close
                '    Set rst3 = Nothing
                '    Record.SubItems(3) = "0"
                '    Record.SubItems(7) = "0.00"
                'Else
                '    Record.SubItems(3) = rst3!Total
                '    Record.SubItems(7) = Format(rst3!TotalApproved, "##,#0.00")
                '    rst3.Close
                '    Set rst3 = Nothing
                'End If
                
                If Len(Record.SubItems(7)) And Len(Record.SubItems(3)) Then
                    If Record.SubItems(7) > "0.00" Or Record.SubItems(3) > "0" Then
                        Record.SubItems(13) = Round((Record.SubItems(7) / Record.SubItems(3)), 2)
                    Else
                        Record.SubItems(13) = "0"
                    End If
                Else
                    Record.SubItems(13) = "0"
                End If
            
                If DIagCode = "" Then
                    strAppend1 = strAppend + " AND (CLNType = '' OR CLNType IS NULL) AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL) "
                    strAppendCase = "11"
                    'sql4 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where (CLNType = '' OR CLNType IS NULL) AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL) " & strAppend & _
                           " GROUP BY CASTreatReason1 "
                Else
                    strAppend1 = strAppend + " AND (CLNType = '' OR CLNType IS NULL) AND CASTreatReason1 = '" & Trim(DIagCode) & "'"
                    strAppendCase = "11"
                    'sql4 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where (CLNType = '' OR CLNType IS NULL) AND CASTreatReason1 = '" & Trim(DIagCode) & "'" & strAppend & _
                           " GROUP BY CASTreatReason1 "
                End If
          
                'rst4.Open sql4, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                Call CalOthersAmt(strAppend1, strAppendCase, Record)
                       
                'If rst4.EOF = True Then
                '    rst4.Close
                '    Set rst4 = Nothing
                '    Record.SubItems(4) = "0"
                '    Record.SubItems(8) = "0.00"
                'Else
                '    Record.SubItems(4) = rst4!Total
                '    Record.SubItems(8) = Format(rst4!TotalApproved, "##,#0.00")
                '    rst4.Close
                '    Set rst4 = Nothing
                'End If
                
                If Record.SubItems(8) > "0.00" Or Record.SubItems(4) > "0" Then
                    Record.SubItems(14) = Round((Record.SubItems(8) / Record.SubItems(4)), 2)
                Else
                    Record.SubItems(14) = "0"
                End If
                
                Record.SubItems(5) = (CDbl(Record.SubItems(2)) + CDbl(Record.SubItems(3)) + CDbl(Record.SubItems(4)))
                Record.SubItems(11) = (CDbl(Record.SubItems(6)) + CDbl(Record.SubItems(7)) + CDbl(Record.SubItems(8)))
                
                If Record.SubItems(11) > "0.00" Or Record.SubItems(5) > "0" Then
                    Record.SubItems(15) = Round((Record.SubItems(11) / Record.SubItems(5)), 2)
                Else
                    Record.SubItems(15) = "0"
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst1.MoveNext
        Loop
        
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub UtilizationByMBM(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMNumber, MBMName, COUNT(CASNumber) AS TotalClaim, " & _
          " SUM(CONVERT(Numeric, MCStatus)) AS MCStatus, SUM(MCDays) AS MCDays, AnnLimit, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY MBMNumber, MBMName, AnnLimit"
        
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "79"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
               
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMNumber) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMNumber)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMName) Then
                    Record.SubItems(1) = Trim(!MBMName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!TotalClaim) Then
                   Record.SubItems(2) = !TotalClaim
                Else
                    Record.SubItems(2) = "0"
                End If
                
                If Len(!MCStatus) Then
                    Record.SubItems(3) = Trim(!MCStatus)
                Else
                    Record.SubItems(3) = "0"
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(4) = Trim(!MCDays)
                Else
                    Record.SubItems(4) = "0"
                End If
                
                If Len(!AnnLimit) Then
                    Record.SubItems(5) = Trim(!AnnLimit)
                Else
                    Record.SubItems(5) = "FULL"
                End If
                
                If Len(!ActualAmt) Then
                    Record.SubItems(6) = Format(!ActualAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                        
                If Len(!ApprovedAmt) Then
                    Record.SubItems(7) = Format(!ApprovedAmt, "#,##0.00")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                
                If !AnnLimit <> "FULL" Then
                    Record.SubItems(8) = Format(Record.SubItems(5) - Record.SubItems(7), "#,##0.00")
                Else
                    Record.SubItems(8) = "FULL"
                End If
                                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub PolGRPCOMListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, MBMPolicyNo, MBMGroupCompany, " & _
          " COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY PAYCode, MBMPolicyNo, MBMGroupCompany "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "88"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    'rst2.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMPolicyNo) Then
                   Record.SubItems(1) = Trim(!MBMPolicyNo)
                End If
                
                If Len(!MBMGroupCompany) Then
                   Record.SubItems(2) = Trim(!MBMGroupCompany)
                End If
                    
                If Len(!TotalCases) Then
                   Record.SubItems(3) = !TotalCases
                End If
                                    
                If Len(!TotalActualAmt) Then
                   Record.SubItems(4) = Format(!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
    'intCol = CInt(lvwList.ColumnHeaders.count - 1)
    lngRow = CLng(lvwList.ListItems.Count - 1)
    
    For X = 0 To CInt(lvwList.ColumnHeaders.Count - 1)
        If ListView2.ColumnHeaders(X + 1).Text <> "" Then
            intCol = intCol + 1
        Else
            Exit For
        End If
    Next
    
    ReDim vntHeader(0)
    ReDim vntData(intCol - 1, lngRow)
    
    'Create Header Array

    For X = 0 To intCol - 1
        ReDim Preserve vntHeader(X)
        vntHeader(X) = lvwList.ColumnHeaders(X + 1).Text
    Next
    
    'Create Data Array

    For X = 0 To lngRow
        vntData(0, X) = lvwList.ListItems.Item(X + 1).Text
        For Y = 1 To intCol - 1
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
    'Select all Cells and and set the number

       'format to string
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
    On Error GoTo debugs
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
    Set objExlSht = Nothing
    Set objExcel = Nothing
Err_OpenExcel:


    Select Case Err
        Case 0
        Case 439
        MsgBox "You must have Microsoft Excel installed on your PC.", vbCritical, "Application Not Found"
        Case Else
        MsgBox Err & ": " & error, vbCritical, "OpenExcel Error"
    End Select
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtActualAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtActualAmt2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAppAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAppAmt2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub cboPanelStatus_Change()
    If InStr(cboPanelStatus.Text, "null") Then
       cboPanelStatus.Enabled = False
    End If
End Sub

Private Sub cboPanelStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPanelStatus.Text, cboPanelStatus.SelStart) + Chr(KeyAscii) + Mid(cboPanelStatus.Text, cboPanelStatus.SelStart + cboPanelStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPanelStatus.ListCount - 1
 
        If NewText = LCase(Left(cboPanelStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPanelStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPanelStatus.Text = ValidValue
                cboPanelStatus.SelStart = 0
                cboPanelStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPayType_Change()
    If InStr(CboPayType.Text, "null") Then
       CboPayType.Enabled = False
    End If
End Sub

Private Sub CboPayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayType.Text, CboPayType.SelStart) + Chr(KeyAscii) + Mid(CboPayType.Text, CboPayType.SelStart + CboPayType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayType.ListCount - 1
 
        If NewText = LCase(Left(CboPayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayType.Text = ValidValue
                CboPayType.SelStart = 0
                CboPayType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClmPending_Change()
    If InStr(CboClmPending.Text, "null") Then
       CboClmPending.Enabled = False
    End If
End Sub

Private Sub CboClmPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClmPending.Text, CboClmPending.SelStart) + Chr(KeyAscii) + Mid(CboClmPending.Text, CboClmPending.SelStart + CboClmPending.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClmPending.ListCount - 1
 
        If NewText = LCase(Left(CboClmPending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClmPending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClmPending.Text = ValidValue
                CboClmPending.SelStart = 0
                CboClmPending.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub MCSummary(Optional strAppend As String)
Const METHOD_NAME = "MCSummary"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
       
    ListView2.ListItems.Clear
    Current = 0
    total = 0
    
    'sql = " SELECT Count(CASNumber) AS NoofMC, MCDays, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql1 = " Group by MCDays "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql1
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "71"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
              
       Do Until rst1.EOF
       
       'total = rst1.RecordCount
       'lblTotalRecord = total
        
                
            With rst1
                    
                    If Len(!MCDays) Then
                        Set Record = ListView2.ListItems.Add(, , !MCDays)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "0")
                    End If
                    
                    If Len(!NoofMC) Then
                       Record.SubItems(1) = Trim(!NoofMC)
                    End If
                    
                    If Len(!ActualAmt) Then
                       Record.SubItems(2) = Format(!ActualAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!AmtClaims) Then
                       Record.SubItems(3) = Format(!AmtClaims, "#,##0.00;(#,##0.00)")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub MCSummary2(Optional strAppend As String)
Const METHOD_NAME = "MCSummary2"
On Error GoTo debugs

Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
        
    ListView2.ListItems.Clear
    Current = 0
    total = 0
    
    'sql = " SELECT ClinicCode, CLNName, Count(CASNumber) AS NoofMC, MCDays, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql1 = " Group by ClinicCode, CLNName, MCDays "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql1
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "72"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
              
       Do Until rst1.EOF
       
       'total = rst1.RecordCount
       'lblTotalRecord = total
        
                
            With rst1
                    
                    If Len(!ClinicCode) Then
                        Set Record = ListView2.ListItems.Add(, , !ClinicCode)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(!CLNName) Then
                       Record.SubItems(1) = Trim(!CLNName)
                    End If
                    
                    If Len(!MCDays) Then
                       Record.SubItems(2) = Trim(!MCDays)
                    Else
                       Record.SubItems(2) = "0"
                    End If
                    
                    If Len(!NoofMC) Then
                       Record.SubItems(3) = Trim(!NoofMC)
                    End If
                    
                    If Len(!ActualAmt) Then
                       Record.SubItems(4) = Format(!ActualAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!AmtClaims) Then
                       Record.SubItems(5) = Format(!AmtClaims, "#,##0.00;(#,##0.00)")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub MCSummary3(Optional strAppend As String)
Const METHOD_NAME = "MCSummary3"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
        
    ListView2.ListItems.Clear
    Current = 0
    total = 0
    
    'sql = " SELECT MBMGroupCompany, Count(CASNumber) AS NoofMC, MCDays, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql1 = " Group by MBMGroupCompany, MCDays "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql1
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "73"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
              
       Do Until rst1.EOF
       
       'total = rst1.RecordCount
       'lblTotalRecord = total
        
                
            With rst1
                    
                    If Len(!MBMGroupCompany) Then
                        Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(!MCDays) Then
                       Record.SubItems(1) = Trim(!MCDays)
                    Else
                       Record.SubItems(1) = "0"
                    End If
                    
                    If Len(!NoofMC) Then
                       Record.SubItems(2) = Trim(!NoofMC)
                    End If
                    
                    If Len(!ActualAmt) Then
                       Record.SubItems(3) = Format(!ActualAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!AmtClaims) Then
                       Record.SubItems(4) = Format(!AmtClaims, "#,##0.00;(#,##0.00)")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub MCSummary4(Optional strAppend As String)
Const METHOD_NAME = "MCSummary4"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim StrSql As String
        
    ListView2.ListItems.Clear
    Current = 0
    total = 0
    DIagCode = ""
    
    'sql = " SELECT CASTreatReason1, Count(CASNumber) AS NoofMC, MCDays, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql1 = " Group by CASTreatReason1, MCDays "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql1
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "74"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
       Do Until rst1.EOF
                
            With rst1
                    
                    If Len(!CASTreatReason1) Then
                        Set Record = ListView2.ListItems.Add(, , !CASTreatReason1)
                        DIagCode = !CASTreatReason1
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                        DIagCode = ""
                    End If
                
                    'sql2 = " SELECT ReasonCode, ReasonDesc From TreatmentReason Where ReasonCode = '" & Trim(DIagCode) & "'"
                    'rst3.Open sql2, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
                    StrSql = "EXEC CISClaimTracking '" & Trim(DIagCode) & "',10"
                    rst2.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                           
                    If rst2.EOF = True Then
                        rst2.Close
                        Set rst2 = Nothing
                        Record.SubItems(1) = ""
                    Else
                        Description = Trim(rst2!ReasonDesc)
                        Record.SubItems(1) = Trim(Description)
                        rst2.Close
                        Set rst2 = Nothing
                    End If
                    
                    If Len(!MCDays) Then
                       Record.SubItems(2) = Trim(!MCDays)
                    Else
                       Record.SubItems(2) = "0"
                    End If
                    
                    If Len(!NoofMC) Then
                       Record.SubItems(3) = Trim(!NoofMC)
                    End If
                    
                    If Len(!ActualAmt) Then
                       Record.SubItems(4) = Format(!ActualAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!AmtClaims) Then
                       Record.SubItems(5) = Format(!AmtClaims, "#,##0.00;(#,##0.00)")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub GRPCOMSumListing(Optional strAppend As String)
Const METHOD_NAME = "GRPCOMSumListing"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, MBMGroupCompany, COUNT(MBMGroupCompany) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY PAYCode, MBMGroupCompany "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "63"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMGroupCompany) Then
                    Record.SubItems(1) = Trim(!MBMGroupCompany)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                
                If Len(!total) Then
                    Record.SubItems(2) = !total
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(5) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(6) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub GRPCOMSumListing2(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMGroupCompany, COUNT(MBMGroupCompany) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " GROUP BY MBMGroupCompany "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "64"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(!MBMGroupCompany))
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!total) Then
                    Record.SubItems(1) = !total
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(2) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(3) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(4) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(5) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub GRPCOMDetailListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT PAYCode, MBMGroupCompany, CASNumber, MBMNumber, CoverID, " & _
          " PatientName, CLMTotalActual, CLMTotalApproved, " & _
          " CLMRecoveryPaidbyM, DateVisit, " & _
          " CASStatus, Claimability, CASTreatReason1, MCDays From VIEW_CLMReport Where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "57"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
    
            Do Until rst1.EOF
            With rst1
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMGroupCompany) Then
                    Record.SubItems(1) = Trim(!MBMGroupCompany)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!CasNumber) Then
                   Record.SubItems(2) = Trim(!CasNumber)
                End If
                                
                If Len(!MBMNumber) Then
                    Record.SubItems(3) = Trim(!MBMNumber)
                End If
                
                If Len(!coverid) Then
                    Record.SubItems(4) = !coverid
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(5) = Trim(!PatientName)
                End If
                
                If Len(!CLMTotalActual) Then
                    Record.SubItems(6) = Format(!CLMTotalActual, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(7) = Format(!CLMTotalApproved, "#,##0.00")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                
                If Len(!CLMRecoveryPaidByM) Then
                    Record.SubItems(8) = Format(!CLMRecoveryPaidByM, "#,##0.00")
                Else
                    Record.SubItems(8) = "0.00"
                End If
                
                If Len(!DateVisit) Then
                    Record.SubItems(9) = Format(!DateVisit, "dd/mm/yyyy")
                End If
                
                If Len(!CASStatus) Then
                   Record.SubItems(10) = !CASStatus
                End If
                
                If Len(!Claimability) Then
                    Record.SubItems(11) = !Claimability
                End If
                                
                If Len(!CASTreatReason1) Then
                    Record.SubItems(12) = Trim(!CASTreatReason1)
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(13) = Trim(!MCDays)
                End If
                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
End Sub

Private Sub GRPCOMEmpUtilizeListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMGroupCompany, MBMNumber, MBMName, MBMCName, CoverID, " & _
        "MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = "GROUP BY MBMDepartment, MBMGroupCompany, MBMNumber, MBMName, MBMCName, CoverID, " & _
        "MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit ORDER BY MBMNumber, MBMGroupCompany"
        
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "58"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMDepartment) Then
                    Record.SubItems(1) = Trim(!MBMDepartment)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(2) = Trim(!MBMNumber)
                End If
                
                If Len(!coverid) Then
                    If !coverid = "00" Then
                        Record.SubItems(3) = Trim(!MBMName)
                    Else
                        If Len(!MBMCName) Then
                            Record.SubItems(3) = Trim(!MBMCName)
                        Else
                            Record.SubItems(3) = "Empty"
                        End If
                    End If
                End If
                                
                If Len(!MBMEmployeeNo) Then
                    Record.SubItems(4) = Trim(!MBMEmployeeNo)
                Else
                    Record.SubItems(4) = "Empty"
                End If
                
                If Len(!AnnLimit) Then
                    Record.SubItems(5) = !AnnLimit
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(6) = Format(!CLMTotalApproved, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                
                If Len(!CLMBalanceLimit) Then
                    Record.SubItems(7) = Format(!CLMBalanceLimit, "#,##0.00")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                
                If Len(!AnnLimit) Then
                    Record.SubItems(8) = Round((!CLMBalanceLimit / !AnnLimit) * 100, 0)
                Else
                    Record.SubItems(8) = 0
                End If
                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub GRPCOMIllnessListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem

    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMGroupCompany, ReasonDesc, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, MBMGroupCompany, ReasonDesc ORDER BY ReasonDesc, MBMGroupCompany"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "60"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMDepartment) Then
                    Record.SubItems(1) = Trim(!MBMDepartment)
                Else
                   Record.SubItems(1) = "Empty"
                End If
                
                If Len(!ReasonDesc) Then
                    Record.SubItems(2) = Trim(!ReasonDesc)
                Else
                   Record.SubItems(2) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(3) = Trim(!Visit)
                Else
                    Record.SubItems(3) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(4) = Trim(!MCTimes)
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(5) = !MCDays
                Else
                    Record.SubItems(5) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(6) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(7) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub GRPCOMClinicListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMGroupCompany, CLNName, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, CLNName, MBMGroupCompany Order By CLNName, MBMGroupCompany"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
    
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "55"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMDepartment) Then
                    Record.SubItems(1) = Trim(!MBMDepartment)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(2) = Trim(!CLNName)
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(3) = Trim(!Visit)
                Else
                    Record.SubItems(3) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(4) = Trim(!MCTimes)
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(5) = !MCDays
                Else
                    Record.SubItems(5) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(6) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(7) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub GRPCOMDeptListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, MBMGroupCompany Order By MBMDepartment, MBMGroupCompany"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "56"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!MBMDepartment) Then
                    Record.SubItems(1) = Trim(!MBMDepartment)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(2) = Trim(!Visit)
                Else
                    Record.SubItems(2) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(3) = Trim(!MCTimes)
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(4) = !MCDays
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(5) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(6) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub GRPCOMMCTimesListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, MBMGroupCompany Order By MCTimes Desc"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "62"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!MBMDepartment) Then
                    Record.SubItems(1) = Trim(!MBMDepartment)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(3) = Trim(!Visit)
                Else
                    Record.SubItems(3) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(2) = Trim(!MCTimes)
                Else
                    Record.SubItems(2) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(4) = !MCDays
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(5) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(6) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub GRPCOMMCDaysListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, MBMGroupCompany Order By MCDays Desc"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "61"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!MBMDepartment) Then
                    Record.SubItems(1) = Trim(!MBMDepartment)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(3) = Trim(!Visit)
                Else
                    Record.SubItems(3) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(4) = Trim(!MCTimes)
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(2) = !MCDays
                Else
                    Record.SubItems(2) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(5) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(6) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub GRPCOMVisitListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, MBMGroupCompany Order By Visit Desc"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "65"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!MBMDepartment) Then
                    Record.SubItems(1) = Trim(!MBMDepartment)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(2) = Trim(!Visit)
                Else
                    Record.SubItems(2) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(3) = Trim(!MCTimes)
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(4) = !MCDays
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(5) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(6) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub GRPCOMGrpCOMListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, MBMGroupCompany Order By MBMGroupCompany, MBMDepartment"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "59"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMGroupCompany) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMGroupCompany)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!MBMDepartment) Then
                    Record.SubItems(1) = Trim(!MBMDepartment)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(2) = Trim(!Visit)
                Else
                    Record.SubItems(2) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(3) = Trim(!MCTimes)
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(4) = !MCDays
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(5) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(6) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub EffMthSummary(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
              
    'sql2 = " GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "47"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Year & "/" & !Month) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = !Year & "/" & Month
                    Else
                        Record.SubItems(1) = !Year & "/" & !Month
                    End If
                End If
                                                
                If Len(!total) Then
                    Record.SubItems(2) = !total
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(3) = !MCDays
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(6) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(7) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub EffMthSummary2(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
   
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
              
    'sql2 = " GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "48"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!Year & "/" & !Month) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                    End If
                End If
                                
                If Len(!total) Then
                    Record.SubItems(1) = !total
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(2) = !MCDays
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(5) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(6) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ExpMthSummary(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
              
    'sql2 = " GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "49"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Year & "/" & !Month) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = !Year & "/" & Month
                    Else
                        Record.SubItems(1) = !Year & "/" & !Month
                    End If
                End If
                                                
                If Len(!total) Then
                    Record.SubItems(2) = !total
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(3) = !MCDays
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(6) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(7) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ExpMthSummary2(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
              
    'sql2 = " GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "50"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!Year & "/" & !Month) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                    End If
                End If
                                
                If Len(!total) Then
                    Record.SubItems(1) = !total
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(2) = !MCDays
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(5) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(6) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub VisitMthSummary(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
               
    'sql = " SELECT PAYCode, MONTH(DateVisit) AS MONTH, YEAR(DateVisit) AS YEAR, COUNT(CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
              
    'sql2 = " GROUP BY PAYCode, MONTH(DateVisit), YEAR(DateVisit) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "51"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Year & "/" & !Month) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = !Year & "/" & Month
                    Else
                        Record.SubItems(1) = !Year & "/" & !Month
                    End If
                End If
                                                
                If Len(!total) Then
                    Record.SubItems(2) = !total
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(3) = !MCDays
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(6) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(7) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub VisitMthSummary2(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MONTH(DateVisit) AS MONTH, YEAR(DateVisit) AS YEAR, COUNT(CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
              
    'sql2 = " GROUP BY MONTH(DateVisit), YEAR(DateVisit) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "52"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!Year & "/" & !Month) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                    End If
                End If
                                
                If Len(!total) Then
                    Record.SubItems(1) = !total
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(2) = !MCDays
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(5) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(6) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub RegMthSummary(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT PAYCode, MONTH(RegisterDate) AS MONTH, YEAR(RegisterDate) AS YEAR, COUNT(CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
              
    'sql2 = " GROUP BY PAYCode, MONTH(RegisterDate), YEAR(RegisterDate) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "53"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Year & "/" & !Month) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = !Year & "/" & Month
                    Else
                        Record.SubItems(1) = !Year & "/" & !Month
                    End If
                End If
                                                
                If Len(!total) Then
                    Record.SubItems(2) = !total
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(3) = !MCDays
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(6) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(7) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub RegMthSummary2(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
                
    'sql = " SELECT MONTH(RegisterDate) AS MONTH, YEAR(RegisterDate) AS YEAR, COUNT(CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(MCDays) AS MCDays From VIEW_CLMReport Where 0 = 0 "
              
    'sql2 = " GROUP BY MONTH(RegisterDate), YEAR(RegisterDate) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "54"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
                
            With rst2
                If Len(!Year & "/" & !Month) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                    End If
                End If
                                
                If Len(!total) Then
                    Record.SubItems(1) = !total
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(2) = !MCDays
                End If
                
                If Len(!TotalActual) Then
                   Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                End If
                                
                If Len(!total) <> 0 Then
                    Record.SubItems(5) = Format(Round(!TotalActual / rst2!total, 2), "#,##0.00")
                    Record.SubItems(6) = Format(Round(!TotalApproved / rst2!total, 2), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub DepartmentByVisit(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
      
    'sql = " SELECT MBMDepartment, COUNT(CASNumber) as TotalCases, SUM(CLMTotalActual) AS totalActual, " & _
          " SUM(CLMTotalApproved) As TotalApproved From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Group By MBMDepartment "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
     
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "45"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(2) = Format(!TotalActual, "#,##0.00")
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(3) = Format(!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub DepartmentByMC(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
        
    ListView2.ListItems.Clear
    Current = 0
    total = 0
    
    'sql = " SELECT MBMDepartment, Count(CASNumber) AS NoofMC, MCDays, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql1 = " Group by MBMDepartment, MCDays "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql1
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "41"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
              
       Do Until rst1.EOF
               
            With rst1
                    
                    If Len(!MBMDepartment) Then
                        Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(!MCDays) Then
                       Record.SubItems(1) = Trim(!MCDays)
                    Else
                       Record.SubItems(1) = "0"
                    End If
                    
                    If Len(!NoofMC) Then
                       Record.SubItems(2) = Trim(!NoofMC)
                    End If
                    
                    If Len(!ActualAmt) Then
                       Record.SubItems(3) = Format(!ActualAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!AmtClaims) Then
                       Record.SubItems(4) = Format(!AmtClaims, "#,##0.00;(#,##0.00)")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
End Sub

Private Sub DepartmentByEmp(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, MBMNumber, MBMName, MBMCName, CoverID, " & _
        "MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = "GROUP BY MBMDepartment, MBMNumber, MBMName, MBMCName, CoverID, " & _
        "MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit ORDER BY MBMNumber, MBMDepartment"

    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "39"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(1) = Trim(!MBMNumber)
                End If
                
                If Len(!coverid) Then
                    If !coverid = "00" Then
                        Record.SubItems(2) = Trim(!MBMName)
                    Else
                        If Len(!MBMCName) Then
                            Record.SubItems(2) = Trim(!MBMCName)
                        Else
                            Record.SubItems(2) = "Empty"
                        End If
                    End If
                End If
                                
                If Len(!MBMEmployeeNo) Then
                    Record.SubItems(3) = Trim(!MBMEmployeeNo)
                Else
                    Record.SubItems(3) = "Empty"
                End If
                
                If Len(!AnnLimit) Then
                    Record.SubItems(4) = !AnnLimit
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(5) = Format(!CLMTotalApproved, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!CLMBalanceLimit) Then
                    Record.SubItems(6) = Format(!CLMBalanceLimit, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                
                If Len(!AnnLimit) Then
                    Record.SubItems(7) = Round((!CLMBalanceLimit / !AnnLimit) * 100, 0)
                Else
                    Record.SubItems(7) = 0
                End If
                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub DepartmentByIllness(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, ReasonDesc, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, ReasonDesc ORDER BY ReasonDesc, MBMDepartment"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "40"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!ReasonDesc) Then
                    Record.SubItems(1) = Trim(!ReasonDesc)
                Else
                   Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(2) = Trim(!Visit)
                Else
                    Record.SubItems(2) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(3) = Trim(!MCTimes)
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(4) = !MCDays
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(5) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(6) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub DepartmentByClinic(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, CLNName, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment, CLNName Order By CLNName, MBMDepartment"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "37"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(1) = Trim(!CLNName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(2) = Trim(!Visit)
                Else
                    Record.SubItems(2) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(3) = Trim(!MCTimes)
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(4) = !MCDays
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(5) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(6) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub DepartmentByDept(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment Order By MBMDepartment"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "38"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                            
                If Len(!Visit) Then
                    Record.SubItems(1) = Trim(!Visit)
                Else
                    Record.SubItems(1) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(2) = Trim(!MCTimes)
                Else
                    Record.SubItems(2) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(3) = !MCDays
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(4) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(5) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub DepartmentByMCTimes(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment Order By MCTimes Desc"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "43"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(2) = Trim(!Visit)
                Else
                    Record.SubItems(2) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(1) = Trim(!MCTimes)
                Else
                    Record.SubItems(1) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(3) = !MCDays
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(4) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(5) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub DepartmentByMCDays(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment Order By MCDays Desc"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "42"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(2) = Trim(!Visit)
                Else
                    Record.SubItems(2) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(3) = Trim(!MCTimes)
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(1) = !MCDays
                Else
                    Record.SubItems(1) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(4) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(5) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub DepartmentByVisit1(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    total = 0
        
    'sql = " SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, " & _
          " SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
                
    'sql2 = " Group By MBMDepartment Order By Visit Desc"
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'End If
            
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "46"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
    Do
        Do Until rst1.EOF
            
            With rst1
                
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Visit) Then
                    Record.SubItems(1) = Trim(!Visit)
                Else
                    Record.SubItems(1) = 0
                End If
                                
                If Len(!MCTimes) Then
                    Record.SubItems(2) = Trim(!MCTimes)
                Else
                    Record.SubItems(2) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(3) = !MCDays
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(4) = Format(!BillAmt, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!AppAmt) Then
                    Record.SubItems(5) = Format(!AppAmt, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst1.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
    Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub ReportMBMGeneralListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMDepartment, MBMNumber, CoverID, PatientName, Count(CASNumber) AS TotalVisit, " & _
          " SUM(MCStatus) as MCTimes, SUM(MCDays) AS MCDays, " & _
          " SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee, CLNType " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY MBMDepartment, MBMNumber, CoverID, PatientName, CLNType "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "4"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(1) = Trim(!MBMNumber)
                End If
                
                If Len(!coverid) Then
                    Record.SubItems(2) = Trim(!coverid)
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(3) = Trim(!PatientName)
                End If
                
                If Len(!TotalVisit) Then
                    Record.SubItems(4) = Trim(!TotalVisit)
                Else
                    Record.SubItems(4) = 0
                End If
                
                If Len(!MCTimes) Then
                    Record.SubItems(5) = Trim(!MCTimes)
                Else
                    Record.SubItems(5) = 0
                End If
                
                If Len(!MCDays) Then
                    Record.SubItems(6) = Trim(!MCDays)
                Else
                    Record.SubItems(6) = 0
                End If
                
                If Len(!CLMTotalActual) Then
                    Record.SubItems(7) = Format(!CLMTotalActual, "#,##0.00")
                Else
                    Record.SubItems(7) = Format(0, "#,##0.00")
                End If
                
                strAppendCase = "108"
                If Len(!MBMDepartment) Then
                    'strAppend1 = strAppend + " AND MBMNumber = '" & !MBMNumber & "' AND CoverID = '" & !CoverID & "'" & _
                    '            " AND PatientName = '" & !PatientName & "' AND MBMDepartment = '" & !MBMDepartment & "'"
                                
                    strAppend1 = strAppend + " AND MBMNumber = '" & !MBMNumber & "' AND CoverID = '" & !coverid & "'" & _
                                " AND MBMDepartment = '" & !MBMDepartment & "'"
                                
                Else
                    'strAppend1 = strAppend + " AND MBMNumber = '" & !MBMNumber & "' AND CoverID = '" & !CoverID & "'" & _
                    '            " AND PatientName = '" & !PatientName & "' AND (MBMDepartment IS Null or MBMDepartment = '')"
                                
                    strAppend1 = strAppend + " AND MBMNumber = '" & !MBMNumber & "' AND CoverID = '" & !coverid & "'" & _
                                " AND (MBMDepartment IS Null or MBMDepartment = '')"
                                
                End If
                Call CalGPSPAmt(strAppend1, strAppendCase, Record)
                                                    
                'If !CLNType = "S" Then
                '    Record.SubItems(8) = "0.00"
                '    Record.SubItems(9) = Format(!CLMTotalActual - IIf(IsNumeric(!CLMNonGPFee) = True, !CLMNonGPFee, 0), "##,#0.00")
                'Else
                '    Record.SubItems(8) = Format(!CLMTotalActual - IIf(IsNumeric(!CLMNonGPFee) = True, !CLMNonGPFee, 0), "##,#0.00")
                '    Record.SubItems(9) = "0.00"
                'End If
                
                'If Len(!ClmNonGPFee) Then
                '    Record.SubItems(10) = Format(!ClmNonGPFee, "##,#0.00")
                'Else
                '    Record.SubItems(10) = "0.00"
                'End If
                
                'If Len(!CLMMCOProcFee) Then
                '    Record.SubItems(11) = Format(!CLMMCOProcFee, "##,#0.00")
                'Else
                '    Record.SubItems(11) = "0.00"
                'End If
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ReportCLMGeneralListing(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
            
        
    'sql = " SELECT MBMDepartment, CASNumber, BordxRefNo, BordxDate, MBMNumber, CoverID, " & _
            " PatientName, ClinicCode, CLNName, INVPayTo, CASTreatReason1, CLNType, " & _
            " CASDiagnosis5, ReasonDesc, CLMTotalActual, CLMNonGpFee, DateVisit, MCDays " & _
            " From VIEW_CLMReport Where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'Else
    '    sql = sql
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "1"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!CasNumber) Then
                    Record.SubItems(1) = Trim(!CasNumber)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!Bordxrefno) Then
                    Record.SubItems(2) = Trim(!Bordxrefno)
                End If
                
                If Len(!BordxDate) Then
                    Record.SubItems(3) = Trim(!BordxDate)
                End If
                
                If Len(!MBMNumber) Then
                    Record.SubItems(4) = Trim(!MBMNumber)
                End If
                
                If Len(!coverid) Then
                    Record.SubItems(5) = Trim(!coverid)
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(6) = Trim(!PatientName)
                End If
                
                If Len(!ClinicCode) Then
                    Record.SubItems(7) = Trim(!ClinicCode)
                End If
                
                If Len(!CLNName) Then
                   Record.SubItems(8) = !CLNName
                End If
                
                If Len(!INVPayTo) Then
                   Record.SubItems(9) = !INVPayTo
                End If
                
                If Len(!CASTreatReason1) Then
                        Record.SubItems(10) = !CASTreatReason1
                        Record.SubItems(11) = IIf(Len(!ReasonDesc), !ReasonDesc, "")
                Else
                        Record.SubItems(10) = "Empty"
                        Record.SubItems(11) = "Empty"
                End If
                
                If Len(!CLMTotalActual) Then    'Bill Amt
                   Record.SubItems(12) = Format(!CLMTotalActual, "##,#0.00")
                End If
                
                If !CLNType = "S" Then
                    Record.SubItems(13) = "0.00"
                    Record.SubItems(14) = Format(!CLMBordxAmt - IIf(IsNumeric(!ClmNonGPFee) = True, !ClmNonGPFee, 0), "##,#0.00")
                Else
                    Record.SubItems(13) = Format(!CLMBordxAmt - IIf(IsNumeric(!ClmNonGPFee) = True, !ClmNonGPFee, 0), "##,#0.00")
                    Record.SubItems(14) = "0.00"
                End If
                                  
                If Len(!ClmNonGPFee) Then   'Insured Amt
                   Record.SubItems(15) = Format(!ClmNonGPFee, "#,##0.00")
                Else
                   Record.SubItems(15) = "0.00"
                End If
                                         
                If Len(!DateVisit) Then   'DOV
                   Record.SubItems(16) = !DateVisit
                End If
                                                           
                If Len(!MCDays) Then   'MCDays
                    Record.SubItems(17) = !MCDays
                Else
                    Record.SubItems(17) = 0
                End If
                    
                If Len(!CASTreatReason2) Then
                    Record.SubItems(18) = !CASTreatReason2
                    Record.SubItems(19) = IIf(Len(!ReasonDesc2), !ReasonDesc2, "")
                Else
                    Record.SubItems(18) = "Empty"
                    Record.SubItems(19) = "Empty"
                End If
                    
                If Len(!CASTreatReason3) Then
                    Record.SubItems(20) = !CASTreatReason3
                    Record.SubItems(21) = IIf(Len(!ReasonDesc3), !ReasonDesc3, "")
                Else
                    Record.SubItems(20) = "Empty"
                    Record.SubItems(21) = "Empty"
                End If
                    
                If Len(!CASTreatReason4) Then
                    Record.SubItems(22) = !CASTreatReason4
                    Record.SubItems(23) = IIf(Len(!ReasonDesc4), !ReasonDesc4, "")
                Else
                    Record.SubItems(22) = "Empty"
                    Record.SubItems(23) = "Empty"
                End If
                    
                If Len(!CASTreatReason5) Then
                    If Trim(!CASTreatReason5) <> "199" And Trim(!CASTreatReason5) <> "299" Then
                        Record.SubItems(24) = !CASTreatReason5
                        Record.SubItems(25) = IIf(Len(!ReasonDesc5), !ReasonDesc5, "")
                    Else
                        Record.SubItems(24) = !CASTreatReason5
                        Record.SubItems(25) = IIf(Len(!CASDiagnosis5), !CASDiagnosis5, "")
                    End If
                Else
                    Record.SubItems(24) = "Empty"
                    Record.SubItems(25) = "Empty"
                End If
                    
            End With
            
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ReportCLNSummary(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
            
        
    'sql = " SELECT ClinicCode, CLNName, CLNType, COUNT(CASNumber) as NoofVisit, " & _
            " SUM(MCStatus) as MCTimes, SUM(MCDays) as MCDays, SUM(CLMTotalActual) as BillAmt " & _
            " From VIEW_CLMReport Where 0 = 0 "
          
    'sql2 = " GROUP BY ClinicCode, CLNName, CLNType "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "2"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
        Do Until rst2.EOF
        
            With rst2
                If Len(!ClinicCode) Then
                    Set Record = ListView2.ListItems.Add(, , !ClinicCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!CLNName) Then
                    Record.SubItems(1) = Trim(!CLNName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(!CLNType) Then
                    Record.SubItems(2) = Trim(!CLNType)
                End If
                
                If Len(!NoofVisit) Then
                    Record.SubItems(3) = Trim(!NoofVisit)
                Else
                    Record.SubItems(3) = 0
                End If
                
                If Len(!MCTimes) And !MCTimes <> 0 Then
                    Record.SubItems(4) = Trim(!MCTimes)
                    Record.SubItems(8) = Format(IIf(Len(!NoofVisit) And !NoofVisit <> 0, Round((!MCTimes / !NoofVisit) * 100, 2), 0))
                Else
                    Record.SubItems(4) = 0
                    Record.SubItems(8) = 0
                End If
                
                If Len(!MCDays) And !MCDays <> 0 Then
                    Record.SubItems(5) = Trim(!MCDays)
                    Record.SubItems(9) = Format(IIf(Len(!MCTimes) And !MCTimes <> 0, !MCDays / !MCTimes, 0), "#,##0.00")
                Else
                    Record.SubItems(5) = 0
                    Record.SubItems(9) = 0
                End If
                
                If Len(!BillAmt) Then
                    Record.SubItems(6) = Format(!BillAmt, "#,##0.00")
                    Record.SubItems(7) = Format(IIf(Len(!NoofVisit) And !NoofVisit <> 0, !BillAmt / !NoofVisit, "0.00"), "#,##0.00")
                Else
                    Record.SubItems(6) = "0.00"
                    Record.SubItems(7) = "0.00"
                End If
                                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub ReportIllnessSummary(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim StrSql As String
Dim TotalBill As String
    
    DIagCode = ""
    TotalBill = 0
    lblCurrent = 0
    total = ""
    
    ListView2.ListItems.Clear
       
    'sql = " SELECT CASTreatReason1, Sum(CLMTotalActual) As CLMTotalActual, " & _
          " SUM(CLMNonGpFee) AS CLMNonGpFee From VIEW_CLMReport Where 0 = 0 " & strAppend & _
          " GROUP BY CASTreatReason1 "
          
    'rst1.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    'rst4.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    
    strAppendCase = "6"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = cmd.Execute
    
    Do Until rst3.EOF
        TotalBill = TotalBill + rst3!CLMTotalActual
        rst3.MoveNext
    Loop
    rst3.Close
    Set rst3 = Nothing
    
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = cmd.Execute
    
        Do Until rst1.EOF
            With rst1
                If Len(!CASTreatReason1) Then
                    Set Record = ListView2.ListItems.Add(, , !CASTreatReason1)
                    DIagCode = !CASTreatReason1
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    DIagCode = ""
                End If
                
                'sql2 = " SELECT ReasonCode, ReasonDesc From TreatmentReason Where ReasonCode = '" & Trim(DIagCode) & "'"
                'rst2.Open sql2, MediNeptuneConnCISDB, adOpenForwardOnly, adLockReadOnly
                StrSql = "EXEC CISClaimTracking '" & Trim(DIagCode) & "', 10"
                rst2.Open StrSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                       
                If rst2.EOF = True Then
                    rst2.Close
                    Set rst2 = Nothing
                    Record.SubItems(1) = ""
                Else
                    Description = Trim(rst2!ReasonDesc)
                    Record.SubItems(1) = Trim(Description)
                    rst2.Close
                    Set rst2 = Nothing
                End If
                                
                If Len(!ClmNonGPFee) Then
                    Record.SubItems(7) = Format(!ClmNonGPFee, "##,#0.00")
                Else
                    Record.SubItems(7) = "0.00"
                End If
                
                If Len(!CLMTotalActual) Then
                    Record.SubItems(8) = Format(!CLMTotalActual, "##,#0.00")
                    Record.SubItems(9) = Round((!CLMTotalActual / TotalBill) * 100, 2)
                Else
                    Record.SubItems(8) = "0.00"
                    Record.SubItems(9) = "0.00"
                End If
                
                If DIagCode = "" Then
                    strAppendCase = "11"
                    strAppend1 = strAppend + " AND CLNType = 'G' AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL)"
                    'sql4 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where CLNType = 'G' AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL) " & strAppend & _
                           " GROUP BY CASTreatReason1 "
                Else
                    strAppendCase = "11"
                    strAppend1 = strAppend + " AND CLNType = 'G' AND CASTreatReason1 = '" & Trim(DIagCode) & "'"
                    'sql4 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where CLNType = 'G' AND CASTreatReason1 = '" & Trim(DIagCode) & "'" & strAppend & _
                           " GROUP BY CASTreatReason1 "
                End If
          
                'rst4.Open sql4, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                Call CalGPAmt(strAppend1, strAppendCase, Record)
                
                If Record.SubItems(5) > "0.00" And Record.SubItems(2) > "0" Then
                    Record.SubItems(10) = Round((Record.SubItems(5) / Record.SubItems(2)), 2)
                Else
                    Record.SubItems(10) = 0
                End If
                
                If DIagCode = "" Then
                    strAppendCase = "11"
                    strAppend1 = strAppend + " AND CLNType = 'S' AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL)"
                    'sql3 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where CLNType = 'S' AND (CASTreatReason1 = '' OR CASTreatReason1 IS NULL) " & strAppend & _
                           " GROUP BY CASTreatReason1 "
                Else
                    strAppendCase = "11"
                    strAppend1 = strAppend + " AND CLNType = 'S' AND CASTreatReason1 = '" & Trim(DIagCode) & "'"
                    'sql3 = " SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved " & _
                           " From VIEW_CLMReport Where CLNType = 'S' AND CASTreatReason1 = '" & Trim(DIagCode) & "'" & strAppend & _
                           " GROUP BY CASTreatReason1 "
                End If
          
                'rst3.Open sql3, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
                Call CalSPAmt(strAppend1, strAppendCase, Record)
                
                If Record.SubItems(6) > "0.00" And Record.SubItems(3) > "0" Then
                    Record.SubItems(11) = Round((Record.SubItems(6) / Record.SubItems(3)), 2)
                Else
                    Record.SubItems(11) = 0
                End If
            
                Record.SubItems(4) = (CDbl(Record.SubItems(2)) + CDbl(Record.SubItems(3)))
                
                If Record.SubItems(8) > "0.00" And Record.SubItems(4) > "0" Then
                    Record.SubItems(12) = Round((Record.SubItems(8) / Record.SubItems(4)), 2)
                Else
                    Record.SubItems(12) = 0
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst1.MoveNext
        Loop
        
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst1.Close
    Set rst1 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub

Private Sub ReportDepartmentalBordx(Optional strAppend As String)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    'sql = " SELECT MBMDepartment, BordxRefNo, BordxDate, CLNType, " & _
            " SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGpFee) as CLMNonGPFee " & _
            " From VIEW_CLMReport Where 0 = 0 "
            
    'sql2 = " GROUP BY MBMDepartment, BordxRefNo, BordxDate, CLNType "
                         
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "3"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
               
        Do Until rst2.EOF
        
            With rst2
                If Len(!MBMDepartment) Then
                    Set Record = ListView2.ListItems.Add(, , !MBMDepartment)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Bordxrefno) Then
                    Record.SubItems(1) = Trim(!Bordxrefno)
                End If
                
                If Len(!BordxDate) Then
                    Record.SubItems(2) = Trim(!BordxDate)
                End If
                
                If Len(!CLMTotalActual) Then
                    Record.SubItems(3) = Format(!CLMTotalActual, "##,#0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                strAppendCase = "107"
                If Len(!MBMDepartment) Then
                    strAppend1 = strAppend + " AND MBMDepartment = '" & !MBMDepartment & "' AND BordxRefNo = '" & !Bordxrefno & "'"
                Else
                    strAppend1 = strAppend + " AND (MBMDepartment IS Null or MBMDepartment = '')AND BordxRefNo = '" & !Bordxrefno & "'"
                End If
                Call CalGPSPAmt(strAppend1, strAppendCase, Record)
                
                If Len(!ClmNonGPFee) Then
                    Record.SubItems(6) = Format(!ClmNonGPFee, "##,#0.00")
                Else
                    Record.SubItems(6) = "0.00"
                End If
                                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
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
End Sub

Private Sub CalGPAmt(strAppend, strAppendCase, Record)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst4 As New ADODB.Recordset

    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst4.CursorLocation = adUseClient
    rst4.CursorType = adOpenForwardOnly
    rst4.CacheSize = 100
    Set rst4 = cmd.Execute
    
    If CboReportType.ListIndex = "15" Then
        If rst4.EOF = True Then
            Record.SubItems(2) = "0"
            Record.SubItems(5) = "0.00"
        Else
            Record.SubItems(2) = rst4!total
            Record.SubItems(5) = Format(rst4!TotalApproved - IIf(IsNumeric(rst4!ClmNonGPFee) = True, rst4!ClmNonGPFee, 0), "##,#0.00")
        End If
        
    ElseIf CboReportType.ListIndex = "21" Then
        If rst4.EOF = True Then
            Record.SubItems(2) = "0"
            Record.SubItems(6) = "0.00"
        Else
            Record.SubItems(2) = rst4!total
            Record.SubItems(6) = Format(rst4!TotalApproved - IIf(IsNumeric(rst4!ClmNonGPFee) = True, rst4!ClmNonGPFee, 0), "##,#0.00")
        End If
    End If
    
    rst4.Close
    Set rst4 = Nothing
    
End Sub

Private Sub CalSPAmt(strAppend, strAppendCase, Record)
Dim cmd As New ADODB.Command
Dim rst3 As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter


    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = cmd.Execute
                       
    If CboReportType.ListIndex = "15" Then
        If rst3.EOF = True Then
            Record.SubItems(3) = "0"
            Record.SubItems(6) = "0.00"
        Else
            Record.SubItems(3) = rst3!total
            Record.SubItems(6) = Format(rst3!TotalApproved - IIf(IsNumeric(rst3!ClmNonGPFee) = True, rst3!ClmNonGPFee, 0), "##,#0.00")
        End If
    ElseIf CboReportType.ListIndex = "21" Then
        If rst3.EOF = True Then
            Record.SubItems(3) = "0"
            Record.SubItems(7) = "0.00"
        Else
            Record.SubItems(3) = rst3!total
            Record.SubItems(7) = Format(rst3!TotalApproved - IIf(IsNumeric(rst3!ClmNonGPFee) = True, rst3!ClmNonGPFee, 0), "##,#0.00")
        End If
    End If
    
    rst3.Close
    Set rst3 = Nothing
    
End Sub

Private Sub CalOthersAmt(strAppend, strAppendCase, Record)
Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter

    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        Record.SubItems(4) = "0"
        Record.SubItems(8) = "0.00"
    Else
        Record.SubItems(4) = rst2!total
        Record.SubItems(8) = Format(rst2!TotalApproved - IIf(IsNumeric(rst2!ClmNonGPFee) = True, rst2!ClmNonGPFee, 0), "##,#0.00")
    End If
    
    rst2.Close
    Set rst2 = Nothing
    
End Sub

Private Sub CalPanelAmt(strAppend, strAppendCase, Record)
Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter

    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If strAppendCase = "12" Then
        Record.SubItems(2) = Format(IIf(IsNumeric(rst2!TotalPanelAmt) = True, rst2!TotalPanelAmt, 0), "#,##0.00")
    ElseIf strAppendCase = "33" Then
        Record.SubItems(3) = Format(IIf(IsNumeric(rst2!TotalNonPanelAmt) = True, rst2!TotalNonPanelAmt, 0), "#,##0.00")
    End If
    
    rst2.Close
    Set rst2 = Nothing
    
End Sub

Private Sub CalCLMAmt(strAppend, strAppendCase, Record)
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim rst2 As New ADODB.Recordset

    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If ClinicType = "" Then
        If strAppendCase = "31" Then
            Record.SubItems(1) = IIf(IsNumeric(rst2!TotalClaimP) = True, rst2!TotalClaimP, 0)
            Record.SubItems(2) = Format(IIf(IsNumeric(rst2!TotalPanelAmt) = True, rst2!TotalPanelAmt, 0), "#,##0.00")
        ElseIf strAppendCase = "32" Then
            Record.SubItems(3) = IIf(IsNumeric(rst2!TotalClaimNP) = True, rst2!TotalClaimNP, 0)
            Record.SubItems(4) = Format(IIf(IsNumeric(rst2!TotalNonPanelAmt) = True, rst2!TotalNonPanelAmt, 0), "#,##0.00")
        End If
    Else
        If strAppendCase = "31" Then
            Record.SubItems(1) = IIf(IsNumeric(rst2!TotalClaimP) = True, rst2!TotalClaimP, 0)
            Record.SubItems(2) = Format(IIf(IsNumeric(rst2!TotalPanelAmt) = True, rst2!TotalPanelAmt, 0), "#,##0.00")
        ElseIf strAppendCase = "32" Then
            Record.SubItems(3) = IIf(IsNumeric(rst2!TotalClaimNP) = True, rst2!TotalClaimNP, 0)
            Record.SubItems(4) = Format(IIf(IsNumeric(rst2!TotalNonPanelAmt) = True, rst2!TotalNonPanelAmt, 0), "#,##0.00")
        End If
    End If
    rst2.Close
    Set rst2 = Nothing

End Sub

Private Sub CalGPSPAmt(strAppend, strAppendCase, Record)
Const METHOD_NAME = "CalGPSPAmt"
On Error GoTo debugs
Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter

    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "CISClaimTracking"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
        
    Do Until rst2.EOF
        If CboReportType.ListIndex = 15 And CboSelection = "Departmental Bordx" Then
            If rst2!CLNType = "S" Then
                Record.SubItems(5) = Format(rst2!CLMTotalActual - IIf(IsNumeric(rst2!ClmNonGPFee) = True, rst2!ClmNonGPFee, 0), "##,#0.00")
            Else
                Record.SubItems(4) = Format(rst2!CLMTotalActual - IIf(IsNumeric(rst2!ClmNonGPFee) = True, rst2!ClmNonGPFee, 0), "##,#0.00")
            End If
        ElseIf CboReportType.ListIndex = 15 And CboSelection = "Utilization (By Member)" Then
            If rst2!CLNType = "S" Then
                Record.SubItems(9) = Format(rst2!CLMTotalActual - IIf(IsNumeric(rst2!ClmNonGPFee) = True, rst2!ClmNonGPFee, 0), "##,#0.00")
            Else
                Record.SubItems(8) = Format(rst2!CLMTotalActual - IIf(IsNumeric(rst2!ClmNonGPFee) = True, rst2!ClmNonGPFee, 0), "##,#0.00")
            End If
        End If
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ListViewHeader24()
Const METHOD_NAME = "ListViewHeader24"
On Error GoTo debugs
    If CboSelection.Text = "Claims Same Date" Then
    
    
        ListView2.ColumnHeaders(1).Text = "Membership No"
        ListView2.ColumnHeaders(2).Text = "Cover ID"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "Patient Name"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Clinic Code"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Clinic Name"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "DOV"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Total Claim"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(8).Text = "Bill Amt"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        
        For AA = 9 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
    
    ElseIf CboSelection.Text = "Claims Same Amount" Then
    
    
        ListView2.ColumnHeaders(1).Text = "Membership No"
        ListView2.ColumnHeaders(2).Text = "Cover ID"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "Patient Name"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Clinic Code"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Clinic Name"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "DOV"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Total Claim"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(8).Text = "Bill Amt"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        
        For AA = 9 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA

    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub



