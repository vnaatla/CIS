VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostingBordxtoGrpGST 
   Caption         =   "Bordx Group Company Export"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   34
      Top             =   7680
      Width           =   11895
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   7080
         TabIndex        =   44
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   8640
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9240
         TabIndex        =   42
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   41
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10200
         TabIndex        =   40
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   7800
         TabIndex        =   39
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdExp 
         Caption         =   "&Post"
         Height          =   330
         Left            =   8520
         TabIndex        =   38
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdSearchCredit 
         Caption         =   "&SearchCR"
         Height          =   330
         Left            =   8160
         TabIndex        =   37
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   330
         Left            =   7560
         TabIndex        =   36
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmdSearchDebit 
         Caption         =   "&SearchDN"
         Height          =   330
         Left            =   9000
         TabIndex        =   35
         Top             =   200
         Width           =   975
      End
   End
   Begin VB.Frame FramePostInfo 
      Height          =   735
      Left            =   80
      TabIndex        =   25
      Top             =   6840
      Width           =   11895
      Begin VB.TextBox LblCrAmt 
         Height          =   285
         Left            =   10680
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox LblDbAmt 
         Height          =   285
         Left            =   9000
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   7080
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox LblExp 
         Height          =   285
         Left            =   5160
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblBatchNo 
         Height          =   285
         Left            =   3120
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboAccPeriod 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   26
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Batch No"
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
         Left            =   2280
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Exp Status"
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
         Left            =   4200
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period"
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
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "CR Amt"
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
         Left            =   10080
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "DR Amt"
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
         Left            =   8280
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
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
         Left            =   6240
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CheckBox chkCheckAll 
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   80
      TabIndex        =   0
      Top             =   260
      Width           =   11895
      Begin VB.CheckBox ChkGrp 
         Height          =   255
         Left            =   11280
         TabIndex        =   7
         Top             =   210
         Width           =   375
      End
      Begin VB.ComboBox cboGRPCompany 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   6
         Top             =   165
         Width           =   3975
      End
      Begin VB.ComboBox cboExpStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   5
         Top             =   425
         Width           =   3975
      End
      Begin VB.TextBox txtBatchNofr 
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
         Left            =   7200
         TabIndex        =   2
         Top             =   710
         Width           =   1695
      End
      Begin VB.TextBox txtBatchNoTo 
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
         Left            =   9360
         TabIndex        =   1
         Top             =   710
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtBordxDatefr 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   285
         Left            =   3720
         TabIndex        =   9
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEXPDateFr 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   405
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEXPDateTo 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Top             =   405
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBordxNoFr 
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
         Left            =   3720
         TabIndex        =   3
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtBordxNoTo 
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
         Left            =   1680
         TabIndex        =   4
         Top             =   660
         Width           =   1455
      End
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "frmPostingBordxtoGrpGST.frx":0000
         Left            =   1680
         List            =   "frmPostingBordxtoGrpGST.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   920
         Width           =   3530
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
         Left            =   120
         TabIndex        =   52
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   11325
         TabIndex        =   22
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Grp Company"
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
         Left            =   6000
         TabIndex        =   21
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Bordx Date from"
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
         TabIndex        =   20
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No. from"
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
         TabIndex        =   19
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Post Status"
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
         Left            =   6000
         TabIndex        =   18
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Post Date from"
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
         TabIndex        =   17
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "To"
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
         TabIndex        =   16
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label8 
         Caption         =   "To"
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
         TabIndex        =   15
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "To"
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
         TabIndex        =   14
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Left            =   9000
         TabIndex        =   13
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Batch No. from"
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
         Left            =   6000
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstBordxGrpCompanyList 
      Height          =   4935
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "chk"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "J/T"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "J/S"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DB"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans.Ref."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AccPeriod"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Trans.Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "AccCode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "D/C"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Clinics"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PLN Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Payor"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "YYMM"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Grp Comp"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Bordx Ref"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Mem No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Batch No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Index"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstBordxGrpCompanyCDList 
      Height          =   4095
      Left            =   1080
      TabIndex        =   33
      Top             =   2760
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "chk"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "J/T"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "J/S"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DB"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans.Ref."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AccPeriod"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Trans.Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "AccCode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "D/C"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Clinics"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PLN Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Payor"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "YYMM"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Grp Comp"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Bordx Ref"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Mem No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "InsType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Batch No"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Bordx to Group Company Export"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPostingBordxtoGrpGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Boolean
Dim PostingType As String
Dim PostingStatus As Boolean
Dim SearchStatus As String

Private Sub cboExpStatus_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(cboExpStatus, KeyAscii)
End Sub

Private Sub cboGrpCompany_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(cboGRPCompany, KeyAscii)
End Sub

Private Sub chkCheckAll_Click()
Const METHOD_NAME = "chkCheckAll_Click"
On Error GoTo debugs

Dim lvCount As Long
Dim lvIndex As Long
Dim lvCountS As Long

    lvCountS = 0
    If lstBordxGrpCompanyList.ListItems.Count > 0 Then
        If chkCheckAll.Value = 1 Then
            If LblExp = "Yes" Then
                MsgBox "Record has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
                lstBordxGrpCompanyList.SelectedItem.Checked = False
                chkCheckAll.Value = 0
                Exit Sub
            Else
        
                lvCount = lstBordxGrpCompanyList.ListItems.Count
                    Do
                        lvCountS = lvCountS + 1
                        lstBordxGrpCompanyList.ListItems(lvCountS).Checked = True
                        lstBordxGrpCompanyList.ListItems(lvCountS).Text = "Post"
            
                        lvIndex = lvIndex + 1
                    Loop Until lvIndex = lvCount
            End If
        Else
            If LblExp <> "Yes" Then
                lvCount = lstBordxGrpCompanyList.ListItems.Count
                    Do
                        lvCountS = lvCountS + 1
                        lstBordxGrpCompanyList.ListItems(lvCountS).Checked = False
                            lstBordxGrpCompanyList.ListItems(lvCountS).Text = ""
                        lvIndex = lvIndex + 1
                    Loop Until lvIndex = lvCount
            End If
        End If
    Else
        chkCheckAll.Value = 0
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ChkGrp_Click()
Const METHOD_NAME = "ChkGrp_Click"
On Error GoTo debugs

Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String
    cboGRPCompany.Clear
    If ChkGrp.Value = 1 Then
            tmpSelField = "SELECT GRPCoCode as tmpCode, GRPCoName as tmpName "
            tmpTable = "GRPCompany"
            tmpSelCriteria = ""
            Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, cboGRPCompany, MediConnCISDB, "SearchCISDBTableRec")
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub LoadComboListing(tmpSelField As String, tmpSelCriteria As String, tmpTableName As String, tmpComboBox As ComboBox, tmpConn As ADODB.Connection, tmpStoredProc As String)
Const METHOD_NAME = "LoadComboListing"
On Error GoTo debugs

Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    tmpComboBox.Clear
    Set cmd.ActiveConnection = tmpConn
    cmd.CommandText = tmpStoredProc
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelField)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, tmpSelCriteria)
    cmd.Parameters.Append Prm3
    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    Do Until TmpRec.EOF
        With TmpRec
        If tmpTableName = "GRPCompany" Then
            tmpComboBox.AddItem !tmpName
        ElseIf tmpTableName = "CLN" Then
            tmpComboBox.AddItem !TmpCode & " - " & !tmpName
        ElseIf tmpTableName = "AccPeriod" Then
            tmpComboBox.AddItem !TmpCode & "/" & !tmpName
        End If
        TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub cmdBack_Click()
    lstBordxGrpCompanyCDList.Visible = False
    lstBordxGrpCompanyList.Visible = True
    CmdExp.Enabled = True
    chkCheckAll.Enabled = True
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    cboExpStatus.ListIndex = 0
    cboGRPCompany.Text = ""
    cboAccPeriod.Text = ""
    lblBatchNo = ""
    LblExp = ""
    lblCurrent = ""
    LblDbAmt = ""
    LblCrAmt = ""
    lstBordxGrpCompanyList.ListItems.Clear
    lstBordxGrpCompanyCDList.ListItems.Clear
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdExp_Click()
Const METHOD_NAME = "CmdExp_Click"
On Err GoTo ErrorSave
Dim AA As Long
Dim StrPVSql As String
Dim TmpPVIndex As String
Dim UpdAsk
Dim BatchNo As String
Dim CompTrans As Boolean
Dim UpdateSuc As Boolean
Dim ItemChecked As Boolean
Dim TmpBordxIndex As String

    ItemChecked = False
    CompTrans = True
    UpdateSuc = False
    For AA = 1 To lstBordxGrpCompanyList.ListItems.Count
        If lstBordxGrpCompanyList.ListItems(AA).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next AA
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
        Exit Sub
    ElseIf LblExp = "Yes" Then
        MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    ElseIf txtBordxDatefr = "__/__/____" Or txtBordxDateTo = "__/__/____" Then
        MsgBox "Please Enter Bordx Date! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtBordxDatefr.SetFocus
        Exit Sub
    ElseIf Trim(cboAccPeriod.Text) = "" Then
        MsgBox "Please Keyin for Accounting Period ! ", vbOKOnly + vbCritical, DialogBoxTitle
        cboAccPeriod.SetFocus
        Exit Sub
        
    Else
        UpdAsk = MsgBox("Do you want to Post record(s) ?", vbYesNo + vbExclamation)
        If UpdAsk = vbYes Then
            Screen.MousePointer = vbHourglass
            MediConnCISWS.beginTrans
            CompTrans = False
            BatchNo = GenSUNPostingBatchNo
            For AA = 1 To lstBordxGrpCompanyList.ListItems.Count
                TmpBordxIndex = lstBordxGrpCompanyList.ListItems(AA).SubItems(19)
                
                If lstBordxGrpCompanyList.ListItems(AA).Checked = True And Trim(TmpBordxIndex) <> "" Then
                    StrPVSql = ""
                    StrPVSql = "Insert Into SUNPosting(TransRef, BordxGrpCompStatus," & _
                                "PostStatus, PostType, PostDate, PostBy, " & _
                                "AccMonth, AccYear, BatchNo) " & _
                                "Values('" & Trim(lstBordxGrpCompanyList.ListItems(AA).SubItems(4)) & "','1','1','" & PostingType & _
                                "','" & Format(Date, "yyyy/mm/dd") & _
                                "','" & Logon & "','" & Mid(cboAccPeriod, 1, 2) & _
                                "','" & Mid(cboAccPeriod, 4, 4) & "','" & BatchNo & "')"
                    MediConnCISWS.Execute StrPVSql
                End If
            Next AA
            MediConnCISWS.commitTrans
            CompTrans = True
            UpdateSuc = True
            lblBatchNo = BatchNo
            
            StrPVSql = ""
            StrPVSql = " and BatchNo = '" & Trim(lblBatchNo) & "'"
            Call GrpCompListing(StrPVSql)
            PostingStatus = True
            Call CmdPrintFile_Click
            MsgBox "Record(s) Posted Sucessfully... ! ", vbOKOnly
            Screen.MousePointer = vbDefault
            cboExpStatus.ListIndex = 1
        End If
    End If


ErrorSave:
            Screen.MousePointer = vbDefault
            If CompTrans = False Then
                medixmasterconnWS.RollbackTrans
            End If
                MediConnCISDB.Close
                Set MediConnCISDB = Nothing
                medixmasterconn.Close
                Set medixmasterconn = Nothing
                MediConnCISWS.Close
                Set MediConnCISWS = Nothing
                If Err.Description <> "" Then
                    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
                End If
End Sub

Private Sub CmdPrintFile_Click()

Dim SVRPath As String
Dim Filename As String

    Screen.MousePointer = vbHourglass
    If lstBordxGrpCompanyList.ListItems.Count <> 0 Or lstBordxGrpCompanyCDList.ListItems.Count <> 0 Then
        If PostingStatus = False Then
            If SearchStatus = "G" Then
                Call ExportListViewtoExcel(lstBordxGrpCompanyList)
            Else
                Call ExportListViewtoExcel(lstBordxGrpCompanyCDList)
            End If
        Else
            SVRPath = "\\Mars\accounts\Account Journal Import\JIS CLN Not Posted\"
            Filename = lblBatchNo & " BI-Bordx " & Format(txtBordxDatefr, "yyyymmdd") & "-" & Mid(txtBordxDateTo, 1, 2)
            Call ExporttoExcelPosting(lstBordxGrpCompanyList, SVRPath, Filename)
            PostingStatus = False
        End If
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs
    Screen.MousePointer = vbHourglass
    CmdExp.Enabled = False
    CmdPrintFile.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdSearch.Enabled = False
    chkCheckAll.Value = 0
    chkCheckAll.Enabled = False
    SearchStatus = "G"
    lstBordxGrpCompanyCDList.Visible = False
    lstBordxGrpCompanyList.Visible = True
    Call SearchRecord
    CmdExp.Enabled = True
    CmdPrintFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    chkCheckAll.Enabled = True
    Screen.MousePointer = Default
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdSearchCredit_Click()
    Screen.MousePointer = vbHourglass
    CmdExp.Enabled = False
    CmdPrintFile.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdSearch.Enabled = False
    chkCheckAll.Value = 0
    chkCheckAll.Enabled = False
    SearchStatus = "C"
    lstBordxGrpCompanyCDList.Visible = True
    lstBordxGrpCompanyList.Visible = False
    Call SearchRecord
    CmdExp.Enabled = True
    CmdPrintFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    Screen.MousePointer = Default

End Sub

Private Sub cmdSearchDebit_Click()
    Screen.MousePointer = vbHourglass
    CmdExp.Enabled = False
    CmdPrintFile.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdSearch.Enabled = False
    chkCheckAll.Value = 0
    chkCheckAll.Enabled = False
    SearchStatus = "D"
    lstBordxGrpCompanyCDList.Visible = True
    lstBordxGrpCompanyList.Visible = False
    Call SearchRecord
    CmdExp.Enabled = True
    CmdPrintFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = True
    CmdExit.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPostingBordxtoGrpGST")
    intExit = False
    PostingType = "BordxGrp"
    Call ComboListing
    cboEntityType.ListIndex = 0
End Sub

Private Sub ComboListing()
Dim Acc As New ADODB.Recordset
Dim Accsql As String

    cboExpStatus.AddItem "0 - No"
    cboExpStatus.AddItem "1 - Yes"
    cboExpStatus.ListIndex = 0
        Accsql = "Select AccMonth, AccYear from AccPeriod "
        Acc.Open Accsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        Do Until Acc.EOF
            cboAccPeriod.AddItem Trim(Acc!AccMonth) & "/" & Trim(Acc!AccYear)
            Acc.MoveNext
        Loop
        Acc.Close
        Set Acc = Nothing
End Sub

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs


Dim StrSql As String
Dim strAppend As String
Dim sql As String
    
    strAppend = ""
        
    If Len(cboGRPCompany) Then
        strAppend = strAppend + " and MBMGroupCompany = '" & cboGRPCompany & "'"
    End If
    
    If Len(txtBordxNoFr) Then
        strAppend = strAppend + " and BordxRefNo >= '" & Trim(txtBordxNoTo) & "'"
    End If
    
    If Len(txtBordxNoTo) Then
        strAppend = strAppend + " and BordxRefNo <= '" & Trim(txtBordxNoFr) & "'"
    End If
    
    If (txtBordxDatefr) <> "__/__/____" And IsDate(txtBordxDatefr) = True _
        And (txtBordxDateTo) <> "__/__/____" And IsDate(txtBordxDateTo) = True Then
        sql = ""
'        sql = " AND BordxDate >= '" & Format(Trim(txtBordxDatefr), "mm/dd/yyyy") & _
'              "' And BordxDate <= '" & Format(Trim(txtBordxDateTo), "mm/dd/yyyy") & "'"
        sql = " AND BordxDate >= '" & Format(Trim(txtBordxDatefr), "YYYY/MM/DD") & _
              "' And BordxDate <= '" & Format(Trim(txtBordxDateTo), "YYYY/MM/DD") & "'"
        strAppend = strAppend + sql
    End If
    
    If (txtEXPDateFr) <> "__/__/____" And IsDate(txtEXPDateFr) = True _
        And (txtEXPDateTo) <> "__/__/____" And IsDate(txtEXPDateTo) = True Then
        sql = ""
'        sql = " AND PostDate >= '" & Format(Trim(txtEXPDateFr), "mm/dd/yyyy") & _
'              "' And PostDate <= '" & Format(Trim(txtEXPDateTo), "mm/dd/yyyy") & "'"
        sql = " AND PostDate >= '" & Format(Trim(txtEXPDateFr), "YYYY/MM/DD") & _
              "' And PostDate <= '" & Format(Trim(txtEXPDateTo), "YYYY/MM/DD") & "'"
        strAppend = strAppend + sql
    End If
    
    If Mid(cboExpStatus, 1, 1) = "0" Then
        strAppend = strAppend + " And (BordxGrpCompStatus IS NULL OR BordxGrpCompStatus = '0')"
        strAppend = strAppend + " And (PostStatus IS NULL OR PostStatus = '0')"
    ElseIf Mid(cboExpStatus, 1, 1) = "1" Then
        strAppend = strAppend + " And BordxGrpCompStatus ='1' And PostStatus ='1'"
    End If
    
    If Len(txtBatchNofr) Then
        strAppend = strAppend + " and BatchNo >= '" & Trim(txtBatchNofr) & "'"
    End If
    
    If Len(txtBatchNoTo) Then
        strAppend = strAppend + " and BatchNo <= '" & Trim(txtBatchNoTo) & "'"
    End If
    
    If cboEntityType.ListIndex = 0 Then
        strAppend = strAppend + " AND  PayCode <>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        strAppend = strAppend + " AND  PayCode ='SJ' "
    End If
    
    If SearchStatus = "G" Then
        Call GrpCompListing(strAppend)
    ElseIf SearchStatus = "C" Then
        Call GrpCompCRListing(strAppend)
    ElseIf SearchStatus = "D" Then
        Call GrpCompDRListing(strAppend)
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Private Sub GrpCompListing(strAppend As String)
Const METHOD_NAME = "GrpCompListing"
On Error GoTo debugs

Dim lstRec As ListItem
Dim GRP As New ADODB.Recordset
Dim GrpSql As String
Dim tempPLN, tempTransRef, AccPeriodMth, AccPeriodYr, TransDate  As String
Dim JournalType, TransDesc, AccCode, CRDB, TempPay, tempProdCat As String
Dim AA, CurrRec As Integer
Dim Amt
Dim cmd As New ADODB.Command
Dim App1 As New ADODB.Parameter
Dim App2 As New ADODB.Parameter
Dim HOS As New ADODB.Recordset
Dim strsqlPay As String
Dim GRPSearch As New ADODB.Recordset
Dim strGRPSearch As String
Dim GRPComCode As String
Dim tmpMCOProcFee As String
Dim RecLoop As Integer
Dim Amount As String
Dim BordxProcMth As String
Dim GrpCompTPA As String
Dim RegMonth As String
    
    tempPLN = ""
    tempTransRef = ""
    AccPeriodMth = ""
    AccPeriodYr = ""
    TransDate = ""
    TransDesc = ""
    AccCode = ""
    CRDB = ""
    TempPay = ""
    tempProdCat = ""
    lstBordxGrpCompanyList.ListItems.Clear
    CurrRec = 0
    LblExp = "No"
    PostingType = "BordxGrp"
    Amount = ""
    GrpCompTPA = ""
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SUN_CIS_Posting"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 1800
    Set App1 = cmd.CreateParameter("RptType", adVarChar, adParamInput, 20, PostingType)
    cmd.Parameters.Append App1
    Set App2 = cmd.CreateParameter("StrCriterial", adVarChar, adParamInput, 1000, strAppend)
    cmd.Parameters.Append App2
    GRP.CursorLocation = adUseClient
    GRP.CursorType = adOpenForwardOnly
    GRP.CacheSize = 100
    Set GRP = cmd.Execute
    
    Do While Not GRP.EOF
        
        CurrRec = CurrRec + 1
        With GRP
            If !BORDXPrintMonth = "JANUARY" Then
                BordxProcMth = "01"
            ElseIf !BORDXPrintMonth = "FEBRUARY" Then
                BordxProcMth = "02"
            ElseIf !BORDXPrintMonth = "MARCH" Then
                BordxProcMth = "03"
            ElseIf !BORDXPrintMonth = "APRIL" Then
                BordxProcMth = "04"
            ElseIf !BORDXPrintMonth = "MAY" Then
                BordxProcMth = "05"
            ElseIf !BORDXPrintMonth = "JUNE" Then
                BordxProcMth = "06"
            ElseIf !BORDXPrintMonth = "JULY" Then
                BordxProcMth = "07"
            ElseIf !BORDXPrintMonth = "AUGUST" Then
                BordxProcMth = "08"
            ElseIf !BORDXPrintMonth = "SEPTEMBER" Then
                BordxProcMth = "09"
            ElseIf !BORDXPrintMonth = "OCTOBER" Then
                BordxProcMth = "10"
            ElseIf !BORDXPrintMonth = "NOVEMBER" Then
                BordxProcMth = "11"
            ElseIf !BORDXPrintMonth = "DECEMBER" Then
                BordxProcMth = "12"
            End If
        

            'tempPLN = !PLNCode
            tmpMCOProcFee = IIf(IsNumeric(!CLMMCOProcFee), Round(!CLMMCOProcFee, 3), 0)
            GrpCompTPA = IIf(Len(!GRPMCOStatus), !GRPMCOStatus, "")
            If !BordxGrpCompStatus = "1" Then
                LblExp = "Yes"
            End If
            JournalType = "BI"
            tempTransRef = !Bordxrefno
            TransDate = !BordxDate
            TransDesc = Year(!BordxDate) & "/" & Month(!BordxDate) & "-" & Trim(IIf(Len(!BordxSubsidary), !BordxSubsidary, !MBMGroupCompany))
            'TransDesc = !MBMGroupCompany
            If !INVPayTo = "C" Then
                If tmpMCOProcFee <> 0 And GrpCompTPA = "F" Then
                    Amt = CDbl(Amt) + Format(CDbl(!CLMBordxAmt), "#,##0.00")
                    Amount = Round(Format(CDbl(!CLMBordxAmt), "#,##0.00"), 2)
                Else
                    Amt = CDbl(Amt) + CDbl(Format(!CLMBordxAmt, "#,##0.00"))
                    Amount = Round(Format(!CLMBordxAmt, "#,##0.00"), 2)
                End If
            Else
                Amt = CDbl(Amt) + CDbl(Format(!CLMBordxAmt, "#,##0.00"))
                Amount = Round(Format(!CLMBordxAmt, "#,##0.00"), 2)
            End If
            CRDB = "D"
            If Trim(!INVPayTo) = "C" Then
                TempPay = !PAYCode
            Else
                TempPay = "1RE"
            End If
            'GRPComCode = IIf(Len(!GRPCoCode), !GRPCoCode, "")
            'AccCode = "65" & IIf(Len(!GRPCoCode), !GRPCoCode, "")
            GRPComCode = !GRPCoCode
            
            If Mid(!Bordxrefno, 1, 3) = "KDC" Then
                AccCode = "65" & !CoGRPCode
            ElseIf Mid(!Bordxrefno, 1, 3) = "KCN" Then
                AccCode = "65" & !CoGRPCode
            Else
                If Len(!PAYAccCode) Then
                    'AccCode = "65" & Mid(!PAYAccCode, 3, 4)
                    AccCode = "65" & !PAYAccCode
                ElseIf Trim(!MBMGroupCompany) = "GAS MALAYSIA BERHAD" Then 'Chua request
                    AccCode = "65A002"
                Else
                    AccCode = "65" & !PAYCode
                End If
            End If
            RecLoop = 2
            'If !INVPayTo = "C" Then
            '    Amount = CDbl(Amt) + Format(CDbl(!CLMPayableByMedix2C) + CDbl(tmpMCOProcFee), "#,##0.00")
            'Else
            '    Amount = Format(!CLMPayableByMedix2M, "#,##0.00")
            'End If
            
            'If tmpMCOProcFee <> 0 And GrpCompTPA = "F" Then
            '    RecLoop = RecLoop + 1
            'End If

            For AA = 1 To RecLoop

                If AA = 2 Then
                'strGRPSearch = ""
                    CRDB = "C"
                    AccCode = "6000"
                    
                    'If tmpMCOProcFee <> 0 And GrpCompTPA = "F" And !INVPayTo = "C" Then
                    '    Amount = Format(Round(!CLMPayableByMedix2c, 2), "#,##0.00")
                    'Else
                    '    Amount = Round(Amount, 2)
                    'End If
                    TransDesc = Year(!BordxDate) & "/" & Month(!BordxDate) & "-" & Trim(!MBMGroupCompany)
                    
                End If
                
                
                If AA = 3 Then
                    AccCode = "1002"
                    Amount = tmpMCOProcFee
                    TransDesc = Year(!BordxDate) & "/" & Month(!BordxDate) & "-" & Trim(!MBMGroupCompany)
                End If
                Set lstRec = lstBordxGrpCompanyList.ListItems.Add(, , IIf(!BordxGrpCompStatus = 1, "Posted", ""))
                lstRec.SubItems(1) = JournalType
                'lstRec.SubItems(2) - Journal Source
                
                lstRec.SubItems(3) = "UHC"
                lstRec.SubItems(4) = IIf(Len(Trim(tempTransRef)), Trim(tempTransRef), "")
                'lstRec.SubItems(5) - Line No
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(TransDate), Format(TransDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Trim(Mid(TransDesc, 1, 225))
                lstRec.SubItems(8) = IIf(Len(AccCode), AccCode, "")
                lstRec.SubItems(9) = Format(IIf(IsNumeric(Amount), Round(Amount, 3), 0), "#,##0.00")
                lstRec.SubItems(10) = Trim(CRDB)
                
                'lstRec.SubItems(11)  - CLN
                'lstRec.SubItems(12) - PLN Code
                lstRec.SubItems(13) = Trim(TempPay)
                
                If Len(!BordxDate) Then
                    lstRec.SubItems(14) = Mid(!BordxDate, 9, 2) & Mid(!BordxDate, 4, 2)
                Else
                    lstRec.SubItems(14) = ""
                End If
                
                lstRec.SubItems(15) = GRPComCode
                lstRec.SubItems(16) = Mid(tempTransRef, 4, 5)
                'lstRec.SubItems(17) = MBM No
                lstRec.SubItems(18) = IIf(Len(Trim(!BatchNo)), !BatchNo, "")
                If AA = 2 Then
                    lstRec.SubItems(19) = "Posted"
                End If
                
            Next AA
        End With
        GRP.MoveNext
        
        DoEvents
        If intExit = True Then
           Exit Do
        End If
    Loop
    intExit = False
    
    GRP.Close
    Set GRP = Nothing
    
    LblDbAmt = Format(Amt, "#,##0.00")
    LblCrAmt = Format(Amt, "#,##0.00")
    lblCurrent = CurrRec

    If CurrRec = 0 Then
        MsgBox "Sorry, No Record Found! ", vbOKOnly + vbCritical
    ElseIf LblExp = "No" Then
        CmdExp.Enabled = True
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub GrpCompCRListing(strAppend As String)
Const METHOD_NAME = "GrpCompCRListing"
On Error GoTo debugs

Dim lstRec As ListItem
Dim GRP As New ADODB.Recordset
Dim GrpSql As String
Dim tempPLN, tempTransRef, AccPeriodMth, AccPeriodYr, TransDate  As String
Dim JournalType, TransDesc, AccCode, CRDB, TempPay, tempProdCat As String
Dim AA, CurrRec As Integer
Dim Amt
Dim cmd As New ADODB.Command
Dim App1 As New ADODB.Parameter
Dim App2 As New ADODB.Parameter
Dim HOS As New ADODB.Recordset
Dim strsqlPay As String
Dim GRPSearch As New ADODB.Recordset
Dim strGRPSearch As String
Dim GRPComCode As String
    tempPLN = ""
    tempTransRef = ""
    AccPeriodMth = ""
    AccPeriodYr = ""
    TransDate = ""
    TransDesc = ""
    AccCode = ""
    CRDB = ""
    TempPay = ""
    tempProdCat = ""
    lstBordxGrpCompanyCDList.ListItems.Clear
    CurrRec = 0
    LblExp = "No"
    PostingType = "BordxGrp"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SUN_CIS_Posting"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set App1 = cmd.CreateParameter("RptType", adVarChar, adParamInput, 20, PostingType)
    cmd.Parameters.Append App1
    Set App2 = cmd.CreateParameter("StrCriterial", adVarChar, adParamInput, 1000, strAppend)
    cmd.Parameters.Append App2
    GRP.CursorLocation = adUseClient
    GRP.CursorType = adOpenForwardOnly
    GRP.CacheSize = 100
    Set GRP = cmd.Execute
    
    Do While Not GRP.EOF
        
        CurrRec = CurrRec + 1
        With GRP
            'tempPLN = !PLNCode
            If !BordxGrpCompStatus = "1" Then
                LblExp = "Yes"
            End If
            JournalType = "BI"
            tempTransRef = !Bordxrefno
            TransDate = !BordxDate
            TransDesc = Year(!BordxDate) & "/" & Month(!BordxDate) & "-" & Trim(!MBMGroupCompany)
            Amt = Amt + Format(CDbl(!CLMTotalActual) - CDbl(!CLMCoPayColl), "#,##0.00")
            CRDB = "C"
            If Trim(!INVPayTo) = "C" Then
                TempPay = !PAYCode
            Else
                TempPay = "1RE"
            End If
            'GRPComCode = !GRPCoCode
            'strGRPSearch = "Select GRPCoCode, CoGRPAcccode,GRPCoName from GRPCompany where GRPCoCode = '" & Trim(GRPComCode) & "'"
            'GRPSearch.Open strGRPSearch, MediConnCISDB, adOpenKeyset, adLockOptimistic
           '
            'If GRPSearch.RecordCount Then
            '    AccCode = GRPSearch!CoGRPAcccode
            'End If
            
            'GRPSearch.Close
            'Set GRPSearch = Nothing
            
            'For AA = 1 To 2
            '    If AA = 2 Then
            '    'strGRPSearch = ""
            '        CRDB = "C"
                    AccCode = "6000"
            '    End If
                Set lstRec = lstBordxGrpCompanyCDList.ListItems.Add(, , IIf(!BordxGrpCompStatus = 1, "Posted", ""))
                
                lstRec.SubItems(1) = JournalType
                'lstRec.SubItems(2) - Journal Source
                
                lstRec.SubItems(3) = "UHC"
                lstRec.SubItems(4) = IIf(Len(Trim(tempTransRef)), Trim(tempTransRef), "")
                'lstRec.SubItems(5) - Line No
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(TransDate), Format(TransDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Trim(Mid(TransDesc, 1, 25))
                lstRec.SubItems(8) = IIf(Len(AccCode), AccCode, "")
                lstRec.SubItems(9) = Format(CDbl(!CLMTotalActual) - CDbl(!CLMCoPayColl), "#,##0.00")
                lstRec.SubItems(10) = Trim(CRDB)
                
                'lstRec.SubItems(11)  - CLN
                'lstRec.SubItems(12) - PLN Code
                lstRec.SubItems(13) = Trim(TempPay)
                'lstRec.SubItems(14) = Trim(YYMM)
                 lstRec.SubItems(15) = GRPComCode
                 lstRec.SubItems(16) = Mid(tempTransRef, 4, 5)
                'lstRec.SubItems(17) = MBM No
                lstRec.SubItems(18) = IIf(Len(Trim(!BatchNo)), !BatchNo, "")
                
            'Next AA
        End With
        GRP.MoveNext
        
        DoEvents
        If intExit = True Then
           Exit Do
        End If
    Loop
    intExit = False
    
    GRP.Close
    Set GRP = Nothing
    
    LblDbAmt = Format(Amt, "#,##0.00")
    LblCrAmt = Format(Amt, "#,##0.00")
    lblCurrent = CurrRec

    If CurrRec = 0 Then
        MsgBox "Sorry, No Record Found! ", vbOKOnly + vbCritical
    ElseIf LblExp = "No" Then
        CmdExp.Enabled = True
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub GrpCompDRListing(strAppend As String)
Const METHOD_NAME = "GrpCompDRListing"
On Error GoTo debugs

Dim lstRec As ListItem
Dim GRP As New ADODB.Recordset
Dim GrpSql As String
Dim tempPLN, tempTransRef, AccPeriodMth, AccPeriodYr, TransDate  As String
Dim JournalType, TransDesc, AccCode, CRDB, TempPay, tempProdCat As String
Dim AA, CurrRec As Integer
Dim Amt
Dim cmd As New ADODB.Command
Dim App1 As New ADODB.Parameter
Dim App2 As New ADODB.Parameter
Dim HOS As New ADODB.Recordset
Dim strsqlPay As String
Dim GRPSearch As New ADODB.Recordset
Dim strGRPSearch As String
Dim GRPComCode As String
    tempPLN = ""
    tempTransRef = ""
    AccPeriodMth = ""
    AccPeriodYr = ""
    TransDate = ""
    TransDesc = ""
    AccCode = ""
    CRDB = ""
    TempPay = ""
    tempProdCat = ""
    lstBordxGrpCompanyCDList.ListItems.Clear
    CurrRec = 0
    LblExp = "No"
    PostingType = "BordxGrp"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SUN_CIS_Posting"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set App1 = cmd.CreateParameter("RptType", adVarChar, adParamInput, 20, PostingType)
    cmd.Parameters.Append App1
    Set App2 = cmd.CreateParameter("StrCriterial", adVarChar, adParamInput, 1000, strAppend)
    cmd.Parameters.Append App2
    GRP.CursorLocation = adUseClient
    GRP.CursorType = adOpenForwardOnly
    GRP.CacheSize = 100
    Set GRP = cmd.Execute
    
    Do While Not GRP.EOF
        
        CurrRec = CurrRec + 1
        With GRP
            'tempPLN = !PLNCode
            If !BordxGrpCompStatus = "1" Then
                LblExp = "Yes"
            End If
            JournalType = "BI"
            tempTransRef = !Bordxrefno
            TransDate = !BordxDate
            TransDesc = Year(!BordxDate) & "/" & Month(!BordxDate) & "-" & Trim(!MBMGroupCompany)
            Amt = Amt + Format(CDbl(!CLMTotalActual) - CDbl(!CLMCoPayColl), "#,##0.00")
            CRDB = "D"
            If Trim(!INVPayTo) = "C" Then
                TempPay = !PAYCode
            Else
                TempPay = "1RE"
            End If
            GRPComCode = !GRPCoCode
            strGRPSearch = "Select GRPCoCode, CoGRPAcccode,GRPCoName from GRPCompany where GRPCoCode = '" & Trim(GRPComCode) & "'"
            GRPSearch.Open strGRPSearch, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            If GRPSearch.RecordCount Then
                AccCode = GRPSearch!CoGRPAcccode
            End If
            
            GRPSearch.Close
            Set GRPSearch = Nothing
            
            'For AA = 1 To 2
            '    If AA = 2 Then
            '    'strGRPSearch = ""
            '        CRDB = "C"
            '        AccCode = "6000"
            '    End If
                Set lstRec = lstBordxGrpCompanyCDList.ListItems.Add(, , IIf(!BordxGrpCompStatus = 1, "Posted", ""))
                
                lstRec.SubItems(1) = JournalType
                'lstRec.SubItems(2) - Journal Source
                
                lstRec.SubItems(3) = "UHC"
                lstRec.SubItems(4) = IIf(Len(Trim(tempTransRef)), Trim(tempTransRef), "")
                'lstRec.SubItems(5) - Line No
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(TransDate), Format(TransDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Trim(Mid(TransDesc, 1, 25))
                lstRec.SubItems(8) = IIf(Len(AccCode), AccCode, "")
                lstRec.SubItems(9) = Format(CDbl(!CLMTotalActual) - CDbl(!CLMCoPayColl), "#,##0.00")
                lstRec.SubItems(10) = Trim(CRDB)
                'lstRec.SubItems(11)  - CLN
                'lstRec.SubItems(12) - PLN Code
                lstRec.SubItems(13) = Trim(TempPay)
                'lstRec.SubItems(14) = Trim(YYMM)
                 lstRec.SubItems(15) = GRPComCode
                 lstRec.SubItems(16) = Mid(tempTransRef, 4, 5)
                'lstRec.SubItems(17) = MBM No
                lstRec.SubItems(18) = IIf(Len(Trim(!BatchNo)), !BatchNo, "")

                
            'Next AA
        End With
        GRP.MoveNext
        
        DoEvents
        If intExit = True Then
           Exit Do
        End If
    Loop
    intExit = False
    
    GRP.Close
    Set GRP = Nothing
    
    LblDbAmt = Format(Amt, "#,##0.00")
    LblCrAmt = Format(Amt, "#,##0.00")
    lblCurrent = CurrRec

    If CurrRec = 0 Then
        MsgBox "Sorry, No Record Found! ", vbOKOnly + vbCritical
    ElseIf LblExp = "No" Then
        CmdExp.Enabled = True
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CboKeyPress(tmpCbo As ComboBox, KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    If tmpCbo.Name <> "cboPostStatus" Then
        NewText = LCase(Left(tmpCbo.Text, tmpCbo.SelStart) + Chr(KeyAscii) + Mid(tmpCbo.Text, tmpCbo.SelStart + tmpCbo.SelLength + 1))
    End If
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To tmpCbo.ListCount - 1
 
        If NewText = LCase(Left(tmpCbo.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = tmpCbo.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            tmpCbo.Text = ValidValue
            tmpCbo.SelStart = 0
            tmpCbo.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub lstBordxGrpCompanyList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim CurrChk As Integer
Dim ModResult As Integer
Dim NextChk1 As Integer
Dim NextChk2 As Integer
Dim lstRec As ListItem
Dim PrevText As String
    NextChk2 = 0
    If LblExp = "Yes" Then
        MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Item.Checked = False
    Else
        CurrChk = Item.Index
        If Item.ListSubItems(10) = "D" Then
            NextChk1 = CurrChk + 1
            If CurrChk + 2 <= lstBordxGrpCompanyList.ListItems.Count Then
               If lstBordxGrpCompanyList.ListItems(CurrChk + 2).SubItems(10) = "C" Then
                    NextChk2 = CurrChk + 2
               End If
            End If
        Else
            NextChk1 = CurrChk - 1
            If lstBordxGrpCompanyList.ListItems(NextChk1).SubItems(10) = "C" Then
               NextChk2 = CurrChk - 2
            Else
                If CurrChk + 1 <= lstBordxGrpCompanyList.ListItems.Count Then
                    If lstBordxGrpCompanyList.ListItems(CurrChk + 1).SubItems(10) = "C" Then
                        NextChk2 = CurrChk + 1
                    End If
                End If
            End If
        End If
        
        PrevText = Item.Text
        If Item.Text = "" Then
            Item.Text = "Post"
            Item.Checked = True
            lstBordxGrpCompanyList.ListItems(NextChk1).Text = "Post"
            lstBordxGrpCompanyList.ListItems(NextChk1).Checked = True
            If NextChk2 > 0 Then
                lstBordxGrpCompanyList.ListItems(NextChk2).Text = "Post"
                lstBordxGrpCompanyList.ListItems(NextChk2).Checked = True
            End If
            
        End If
        If PrevText = "Post" Then
            Item.Text = ""
            Item.Checked = False
            lstBordxGrpCompanyList.ListItems(NextChk1).Text = ""
            lstBordxGrpCompanyList.ListItems(NextChk1).Checked = False
            If NextChk2 > 0 Then
                lstBordxGrpCompanyList.ListItems(NextChk2).Text = ""
                lstBordxGrpCompanyList.ListItems(NextChk2).Checked = False
            End If
        End If
    End If
End Sub

Private Sub txtEXPDateFr_LostFocus()
    If IsDate(txtEXPDateFr) Then
    Else
        If txtEXPDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
             txtEXPDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtEXPDateTo_LostFocus()
    If IsDate(txtEXPDateTo) Then
    Else
        If txtEXPDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
             txtEXPDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDateFr_LostFocus()
    If IsDate(txtBordxDatefr) Then
    Else
        If txtBordxDatefr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDatefr.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDateTo_LostFocus()
    If IsDate(txtBordxDateTo) Then
    Else
        If txtBordxDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateTo.SetFocus
        End If
    End If
End Sub








