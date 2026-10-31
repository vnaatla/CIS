VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostingRecGrpCompany 
   Caption         =   "Receipt Group Company Export"
   ClientHeight    =   8364
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   12036
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8364
   ScaleWidth      =   12036
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   80
      TabIndex        =   16
      Top             =   260
      Width           =   11895
      Begin VB.CheckBox ChkGrp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   7
         Top             =   220
         Width           =   255
      End
      Begin VB.ComboBox cboGRPCompany 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtReceiptDatefr 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   165
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtReceiptDateTo 
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Top             =   165
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEXPDateFr 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   405
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEXPDateTo 
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Top             =   405
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRECNoTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtRECNoFr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   660
         Width           =   1455
      End
      Begin VB.ComboBox cboExpStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   8
         Top             =   430
         Width           =   4215
      End
      Begin VB.TextBox txtBatchNofr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtBatchNoTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9600
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "frmPostingRecGrpCompany.frx":0000
         Left            =   1800
         List            =   "frmPostingRecGrpCompany.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   910
         Width           =   3495
      End
      Begin VB.Label lblEntityType 
         Caption         =   "Entity Type"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   11565
         TabIndex        =   46
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "Batch No. from"
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   25
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   420
         Width           =   345
      End
      Begin VB.Label Label6 
         Caption         =   "Post Date from"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Post Status"
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Receipt No. from"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt Date from"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   6000
         TabIndex        =   17
         Top             =   200
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lstRecGrpCompanyList 
      Height          =   4935
      Left            =   75
      TabIndex        =   45
      Top             =   1920
      Width           =   11895
      _ExtentX        =   20976
      _ExtentY        =   8700
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
         Size            =   8.4
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Index"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstRecGrpCompanyCDList 
      Height          =   4935
      Left            =   80
      TabIndex        =   15
      Top             =   1920
      Width           =   11895
      _ExtentX        =   20976
      _ExtentY        =   8700
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
         Text            =   "Ins Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Batch No"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame FramePostInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   80
      TabIndex        =   27
      Top             =   6840
      Width           =   11895
      Begin VB.TextBox LblCrAmt 
         Height          =   285
         Left            =   10680
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox LblDbAmt 
         Height          =   285
         Left            =   9000
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   7080
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox LblExp 
         Height          =   285
         Left            =   5160
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblBatchNo 
         Height          =   285
         Left            =   3000
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboAccPeriod 
         Height          =   315
         ItemData        =   "frmPostingRecGrpCompany.frx":002B
         Left            =   1080
         List            =   "frmPostingRecGrpCompany.frx":002D
         TabIndex        =   11
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   6240
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "DR Amt"
         Height          =   255
         Left            =   8280
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "CR Amt"
         Height          =   255
         Left            =   10080
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Post Status"
         Height          =   255
         Left            =   4200
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CheckBox chkCheckAll 
      Caption         =   "Check All"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   7680
      Width           =   11895
      Begin VB.CommandButton CmdExp 
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         TabIndex        =   14
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   40
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10200
         TabIndex        =   39
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10920
         TabIndex        =   38
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9240
         TabIndex        =   37
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8640
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   12
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdchkAll 
         Caption         =   "&Check All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   35
         Top             =   200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdSearchDebit 
         Caption         =   "&SearchDN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7680
         TabIndex        =   43
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdSearchCredit 
         Caption         =   "&SearchCR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7680
         TabIndex        =   41
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7680
         TabIndex        =   42
         Top             =   200
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Receipt Group Company Export"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmPostingRecGrpCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Boolean
Dim PostingType As String
Dim PostingStatus As Boolean
Dim SearchStatus As String

Private Sub cboAccPeriod_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(cboAccPeriod, KeyAscii)
End Sub

Private Sub cboExpStatus_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(cboExpStatus, KeyAscii)
End Sub


Private Sub cboGRPCompany_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(cboGRPCompany, KeyAscii)
End Sub

Private Sub chkCheckAll_Click()
Dim lvCount As Long
Dim lvIndex As Long
Dim lvCountS As Long

    lvCountS = 0
    If lstRecGrpCompanyList.ListItems.Count > 0 Then
        If chkCheckAll.Value = 1 Then
            If LblExp = "Yes" Then
                MsgBox "Record has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
                lstRecGrpCompanyList.SelectedItem.Checked = False
                chkCheckAll.Value = 0
                Exit Sub
            Else
        
                lvCount = lstRecGrpCompanyList.ListItems.Count
                    Do
                        lvCountS = lvCountS + 1
                        lstRecGrpCompanyList.ListItems(lvCountS).Checked = True
                        lstRecGrpCompanyList.ListItems(lvCountS).Text = "Post"
            
                        lvIndex = lvIndex + 1
                    Loop Until lvIndex = lvCount
            End If
        Else
            If LblExp <> "Yes" Then
                lvCount = lstRecGrpCompanyList.ListItems.Count
                    Do
                        lvCountS = lvCountS + 1
                        lstRecGrpCompanyList.ListItems(lvCountS).Checked = False
                            lstRecGrpCompanyList.ListItems(lvCountS).Text = ""
                        lvIndex = lvIndex + 1
                    Loop Until lvIndex = lvCount
            End If
        End If
    Else
        chkCheckAll.Value = 0
        'MsgBox "No record found on Listing! ", vbOKOnly + vbCritical, DialogBoxTitle
    End If

End Sub

Private Sub ChkGrp_Click()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String

    cboGRPCompany.Clear
    If ChkGrp.Value = 1 Then
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            tmpSelField = "SELECT GRPCoCode as tmpCode, GRPCoName as tmpName "
            tmpTable = "GRPCompany"
            tmpSelCriteria = ""
            Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, cboGRPCompany, MediConnCISDB, "SearchCISDBTableRec")
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub LoadComboListing(tmpSelField As String, tmpSelCriteria As String, tmpTableName As String, tmpComboBox As ComboBox, tmpConn As ADODB.Connection, tmpStoredProc As String)

Const METHOD_NAME = "LoadComboListing"
On Error GoTo debugs

Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    tmpComboBox.Clear
    Set cmd.ActiveConnection = tmpConn
    cmd.CommandText = tmpStoredProc
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelField)
    cmd.Parameters.Append prm1
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
    lstRecGrpCompanyCDList.Visible = False
    lstRecGrpCompanyList.Visible = True
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
    lstRecGrpCompanyList.ListItems.Clear
    lstRecGrpCompanyCDList.ListItems.Clear
End Sub

Private Sub cmdExit_Click()
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
Dim tmpIndex As String

    ItemChecked = False
    CompTrans = True
    UpdateSuc = False
    For AA = 1 To lstRecGrpCompanyList.ListItems.Count
        If lstRecGrpCompanyList.ListItems(AA).Checked = True Then
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
        
    ElseIf txtReceiptDatefr = "__/__/____" Or txtReceiptDateTo = "__/__/____" Then
        MsgBox "Please Enter Receipt Date! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDatefr.SetFocus
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
            For AA = 1 To lstRecGrpCompanyList.ListItems.Count
                tmpIndex = Trim(lstRecGrpCompanyList.ListItems(AA).SubItems(19))
                If lstRecGrpCompanyList.ListItems(AA).Checked = True And tmpIndex <> "" Then
                    StrPVSql = ""
                    StrPVSql = "Insert Into SUNPosting(TransRef, RecGrpCompStatus," & _
                                "PostStatus, PostType, PostDate, PostBy, " & _
                                "AccMonth, AccYear, BatchNo) " & _
                                "Values('" & Trim(lstRecGrpCompanyList.ListItems(AA).SubItems(4)) & "','1','1','" & PostingType & _
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
            
            If UpdateSuc = True Then
                StrPVSql = ""
                StrPVSql = " and BatchNo = '" & Trim(lblBatchNo) & "'"
                Call GrpCompListing(StrPVSql)
                PostingStatus = True
                Call CmdPrintFile_Click
            End If
            
            Screen.MousePointer = vbDefault
            MsgBox "Record(s) Posted Sucessfully... ! ", vbOKOnly
            cboExpStatus.ListIndex = 1
            
        End If
    End If
ErrorSave:
            Screen.MousePointer = vbDefault
            If CompTrans = False Then
                medixmasterconnWS.RollbackTrans
                If Err.Description <> "" Then
                    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
                End If
            End If
            
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub CmdPrintFile_Click()
Dim SVRPath As String
Dim Filename As String

    Screen.MousePointer = vbHourglass
    If lstRecGrpCompanyList.ListItems.Count <> 0 Or lstRecGrpCompanyCDList.ListItems.Count <> 0 Then
        If PostingStatus = False Then
            If SearchStatus = "G" Then
                Call ExportListViewtoExcel(lstRecGrpCompanyList)
            Else
                Call ExportListViewtoExcel(lstRecGrpCompanyCDList)
            End If
        Else
            SVRPath = "\\Mars\accounts\Account Journal Import\JIS CLN Not Posted\"
            Filename = lblBatchNo & " CR-GrpCompany " & Format(txtReceiptDatefr, "yyyymmdd") & "-" & Mid(txtReceiptDateTo, 1, 2)
            Call ExporttoExcelPosting(lstRecGrpCompanyList, SVRPath, Filename)
            PostingStatus = False
        End If
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
    CmdExp.Enabled = False
    CmdPrintFile.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdSearch.Enabled = False
    chkCheckAll.Value = 0
    chkCheckAll.Enabled = False
    SearchStatus = "G"
    lstRecGrpCompanyCDList.Visible = False
    lstRecGrpCompanyList.Visible = True
        Call SearchRecord
    CmdExp.Enabled = True
    CmdPrintFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    chkCheckAll.Enabled = True
    Screen.MousePointer = Default
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
    lstRecGrpCompanyCDList.Visible = True
    lstRecGrpCompanyList.Visible = False
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
    lstRecGrpCompanyCDList.Visible = True
    lstRecGrpCompanyList.Visible = False
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
Call ActionLogUpdate(12, "frmPostingRecGrpCompany")
    intExit = False
    PostingType = "RecGrp"
    Call ComboListing
    cboEntityType.ListIndex = 0
End Sub

Private Sub ComboListing()
Dim GrpComRec As New ADODB.Recordset
Dim GrpComSql As String
Dim Acc As New ADODB.Recordset
Dim Accsql As String

    cboExpStatus.AddItem "0 - No"
    cboExpStatus.AddItem "1 - Yes"
    cboExpStatus.ListIndex = 0
    
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Accsql = "Select AccMonth, AccYear from AccPeriod "
        Acc.Open Accsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        Do Until Acc.EOF
            cboAccPeriod.AddItem Trim(Acc!AccMonth) & "/" & Trim(Acc!AccYear)
            Acc.MoveNext
        Loop
        Acc.Close
        Set Acc = Nothing
 '   medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing
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
    
    If Len(txtRECNoFr) Then
        strAppend = strAppend + " and ReceiptNo >= '" & Trim(txtRECNoFr) & "'"
    End If
    
    If Len(txtRECNoTo) Then
        strAppend = strAppend + " and ReceiptNo <= '" & Trim(txtRECNoTo) & "'"
    End If
    
    If (txtReceiptDatefr) <> "__/__/____" And IsDate(txtReceiptDatefr) = True _
        And (txtReceiptDateTo) <> "__/__/____" And IsDate(txtReceiptDateTo) = True Then
        sql = ""
        sql = " AND ReceiptDate >= '" & Format(Trim(txtReceiptDatefr), "mm/dd/yyyy") & _
              "' And ReceiptDate <= '" & Format(Trim(txtReceiptDateTo), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If (txtEXPDateFr) <> "__/__/____" And IsDate(txtEXPDateFr) = True _
        And (txtEXPDateTo) <> "__/__/____" And IsDate(txtEXPDateTo) = True Then
        sql = ""
        sql = " AND PostDate >= '" & Format(Trim(txtEXPDateFr), "mm/dd/yyyy") & _
              "' And PostDate <= '" & Format(Trim(txtEXPDateTo), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Mid(cboExpStatus, 1, 1) = "0" Then
        strAppend = strAppend + " And (RecGrpCompStatus IS NULL OR RecGrpCompStatus = '0')"
        strAppend = strAppend + " And (PostStatus IS NULL OR PostStatus = '0')"
    ElseIf Mid(cboExpStatus, 1, 1) = "1" Then
        strAppend = strAppend + " And RecGrpCompStatus ='1' And PostStatus ='1'"
    End If
    
    If Len(txtBatchNofr) Then
        strAppend = strAppend + " and BatchNo >= '" & Trim(txtBatchNofr) & "'"
    End If
    
    If Len(txtBatchNoTo) Then
        strAppend = strAppend + " and BatchNo <= '" & Trim(txtBatchNoTo) & "'"
    End If
    
    If cboEntityType.ListIndex = 0 Then
        strAppend = strAppend + " AND  PayCode<>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        strAppend = strAppend + " AND  PayCode='SJ' "
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
Dim GRPSql As String
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
    lstRecGrpCompanyList.ListItems.Clear
    CurrRec = 0
    LblExp = "No"
    PostingType = "RecGrp"
    
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SUN_CIS_Posting"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 1200
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
            If GRP!RecGrpCompStatus = "1" Then
                LblExp = "Yes"
            End If
            JournalType = "CR"
            tempTransRef = !ReceiptNo
            TransDate = !ReceiptDate
            TransDesc = !MBMGroupCompany
            'BordxRefNo
            
            If Mid(!Bordxrefno, 1, 3) = "KDC" Then
                AccCode = "65" & !GRPCoCode
            Else
                If Len(!PAYAccCode) Then
                 AccCode = "65" & Mid(!PAYAccCode, 3, 4)
                Else
                 AccCode = "7610"
                End If
            End If
     
            Amt = Amt + IIf(IsNumeric(!CheckAmt), !CheckAmt, 0)
            CRDB = "D"
            TempPay = !PAYCode
            GRPComCode = !GRPCoCode
            For AA = 1 To 2
            
                If AA = 2 Then
                    'AccCode = "65" & !GRPCoCode
                    CRDB = "C"
                End If
                
                Set lstRec = lstRecGrpCompanyList.ListItems.Add(, , IIf(!RecGrpCompStatus = 1, "Posted", ""))
                
                lstRec.SubItems(1) = JournalType
                lstRec.SubItems(3) = "UHC"
                lstRec.SubItems(4) = IIf(Len(Trim(tempTransRef)), Trim(tempTransRef), "")
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(TransDate), Format(TransDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Trim(TransDesc)
                lstRec.SubItems(8) = IIf(Len(AccCode), AccCode, "")
                lstRec.SubItems(9) = Format(IIf(IsNumeric(!CheckAmt), !CheckAmt, 0), "#,##0.00")
                lstRec.SubItems(10) = Trim(CRDB)
                lstRec.SubItems(13) = Trim(TempPay)
                lstRec.SubItems(15) = GRPComCode
                lstRec.SubItems(16) = !Bordxrefno
                
                lstRec.SubItems(18) = IIf(Len(Trim(!BatchNo)), !BatchNo, "")
                If AA = 2 Then
                    lstRec.SubItems(19) = "Post"
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
Dim GRPSql As String
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
    lstRecGrpCompanyCDList.ListItems.Clear
    CurrRec = 0
    LblExp = "No"
    PostingType = "RecGrp"
    
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
            If GRP!RecGrpCompStatus = "1" Then
                LblExp = "Yes"
            End If
            JournalType = "CR"
            tempTransRef = !ReceiptNo
            TransDate = !ReceiptDate
            TransDesc = !MBMGroupCompany
            AccCode = "7610"
            Amt = Amt + IIf(IsNumeric(!CheckAmt), !CheckAmt, 0)
            CRDB = "C"
            TempPay = !PAYCode
            GRPComCode = !GRPCoCode
            'For AA = 1 To 2
            '    If AA = 2 Then
                'strGRPSearch = ""
                'strGRPSearch = "Select CoGRPAcccode,GRPCoName from GRPCompany where GRPCoName = '" & Trim(TransDesc) & "'"
                'GRPSearch.Open strGRPSearch, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    
                'If GRPSearch.RecordCount Then
                    AccCode = "65" & !GRPCoCode
                'End If
                
                'GRPSearch.Close
                'Set GRPSearch = Nothing
    
             '       CRDB = "C"
             '   End If
                Set lstRec = lstRecGrpCompanyCDList.ListItems.Add(, , IIf(!RecGrpCompStatus = 1, "Posted", ""))
                
                lstRec.SubItems(1) = JournalType
                'lstRec.SubItems(2) - Journal Source
                
                lstRec.SubItems(3) = "UHC"
                lstRec.SubItems(4) = IIf(Len(Trim(tempTransRef)), Trim(tempTransRef), "")
                'lstRec.SubItems(5) - Line No
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(TransDate), Format(TransDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Trim(TransDesc)
                lstRec.SubItems(8) = IIf(Len(AccCode), AccCode, "")
                lstRec.SubItems(9) = Format(IIf(IsNumeric(!CheckAmt), !CheckAmt, 0), "#,##0.00")
                lstRec.SubItems(10) = Trim(CRDB)
                'lstRec.SubItems(11) - CLN
                'lstRec.SubItems(12) - PLN Code
                lstRec.SubItems(13) = Trim(TempPay)
                'lstRec.SubItems(14) = YYMM
                lstRec.SubItems(15) = GRPComCode
                'lstRec.SubItems(16)  - Bordx Ref
                'lstRec.SubItems(17) - MBM No
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
Dim GRPSql As String
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
    lstRecGrpCompanyCDList.ListItems.Clear
    CurrRec = 0
    LblExp = "No"
    PostingType = "RecGrp"
    
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
            If GRP!RecGrpCompStatus = "1" Then
                LblExp = "Yes"
            End If
            JournalType = "CR"
            tempTransRef = !ReceiptNo
            TransDate = !ReceiptDate
            TransDesc = !MBMGroupCompany
            AccCode = "7610"
            Amt = Amt + IIf(IsNumeric(!CheckAmt), !CheckAmt, 0)
            CRDB = "D"
            TempPay = !PAYCode
            GRPComCode = !GRPCoCode
                Set lstRec = lstRecGrpCompanyCDList.ListItems.Add(, , IIf(!RecGrpCompStatus = 1, "Posted", ""))
                
                lstRec.SubItems(1) = JournalType
                lstRec.SubItems(3) = "UHC"
                lstRec.SubItems(4) = IIf(Len(Trim(tempTransRef)), Trim(tempTransRef), "")
                'lstRec.SubItems(5) - Line No
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(TransDate), Format(TransDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Trim(TransDesc)
                lstRec.SubItems(8) = IIf(Len(AccCode), AccCode, "")
                lstRec.SubItems(9) = Format(IIf(IsNumeric(!CheckAmt), !CheckAmt, 0), "#,##0.00")
                lstRec.SubItems(10) = Trim(CRDB)
                'lstRec.SubItems(11) - CLN
                'lstRec.SubItems(12) - PLN Code
                lstRec.SubItems(13) = Trim(TempPay)
                'lstRec.SubItems(14) = YYMM
                lstRec.SubItems(15) = GRPComCode
                'lstRec.SubItems(16)  - Bordx Ref
                'lstRec.SubItems(17) - MBM No
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




Private Sub lstRecGrpCompanyList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim FirstChk As Integer
Dim SecChk As Integer
Dim SEQNumber As Integer
Dim PostStatus As String
Dim lstRecPay As ListItem

    If LblExp = "Yes" Then
        MsgBox "Record has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Item.Checked = False
    Else
    
         FirstChk = Item.Index
         SEQNumber = FirstChk Mod 2
         If SEQNumber = 0 Then
             SecChk = FirstChk - 1
         Else
             SecChk = FirstChk + 1
         End If
    
         PostStatus = Item.Text
         If Item.Text = "" Then
             Item.Text = "Post"
             lstRecGrpCompanyList.ListItems(SecChk).Text = "Post"
             Item.Checked = True
             lstRecGrpCompanyList.ListItems(SecChk).Checked = True
         End If
         If PostStatus = "Post" Then
             Item.Text = ""
             lstRecGrpCompanyList.ListItems(SecChk).Text = ""
             Item.Checked = False
             lstRecGrpCompanyList.ListItems(SecChk).Checked = False
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

Private Sub txtReceiptDatefr_LostFocus()
    If IsDate(txtReceiptDatefr) Then
    Else
        If txtReceiptDatefr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDatefr.SetFocus
        End If
    End If
End Sub

Private Sub txtReceiptDateTo_LostFocus()
    If IsDate(txtReceiptDateTo) Then
    Else
        If txtReceiptDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDateTo.SetFocus
        End If
    End If
End Sub


