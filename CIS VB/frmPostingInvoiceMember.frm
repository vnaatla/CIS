VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostingInvoiceMember 
   Caption         =   "Posting - Member Invoice"
   ClientHeight    =   8304
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11892
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8304
   ScaleWidth      =   11892
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheckAll 
      Caption         =   "Check All"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame FramePostInfo 
      Enabled         =   0   'False
      Height          =   735
      Left            =   60
      TabIndex        =   41
      Top             =   6720
      Width           =   11775
      Begin VB.TextBox LblCrAmt 
         Height          =   285
         Left            =   10560
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox LblDbAmt 
         Height          =   285
         Left            =   8760
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   6960
         TabIndex        =   59
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox LblPost 
         Height          =   285
         Left            =   4920
         TabIndex        =   58
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblBatchNo 
         Height          =   285
         Left            =   2880
         TabIndex        =   57
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboAccPeriod 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   6120
         TabIndex        =   47
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "DN Amt"
         Height          =   255
         Left            =   8040
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "CR Amt"
         Height          =   255
         Left            =   9960
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Post Status"
         Height          =   255
         Left            =   3960
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   2160
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   38
      Top             =   7560
      Width           =   11775
      Begin VB.CommandButton CmdPost 
         Caption         =   "&Post"
         Height          =   330
         Left            =   8520
         TabIndex        =   22
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   7800
         TabIndex        =   21
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10200
         TabIndex        =   24
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   25
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9240
         TabIndex        =   23
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   8640
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   7080
         TabIndex        =   20
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdChkAll 
         Caption         =   "Chk All"
         Enabled         =   0   'False
         Height          =   330
         Left            =   7080
         TabIndex        =   39
         Top             =   200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdSearchDebit 
         Caption         =   "&SearchDN"
         Height          =   330
         Left            =   8160
         TabIndex        =   19
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdSearchCredit 
         Caption         =   "&SearchCR"
         Height          =   330
         Left            =   8160
         TabIndex        =   18
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   330
         Left            =   8280
         TabIndex        =   17
         Top             =   200
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1650
      Left            =   120
      TabIndex        =   26
      Top             =   260
      Width           =   11655
      Begin VB.CheckBox ChkGrp 
         Height          =   255
         Left            =   11115
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   7440
         TabIndex        =   10
         Top             =   195
         Width           =   3615
      End
      Begin VB.ComboBox cboInvReg 
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
         Left            =   7440
         TabIndex        =   12
         Top             =   480
         Width           =   3615
      End
      Begin VB.ComboBox cboPostStatus 
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
         Left            =   7440
         TabIndex        =   13
         Top             =   750
         Width           =   3615
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   195
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   572
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
         Left            =   3960
         TabIndex        =   1
         Top             =   195
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   572
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
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   480
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
      Begin MSMask.MaskEdBox txtPostDate2 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   750
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   480
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
      Begin MSMask.MaskEdBox txtPostDate1 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   750
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox txtBordxNoFr 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   1040
         Width           =   1455
      End
      Begin VB.TextBox txtBordxNoTo 
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
         Left            =   3960
         TabIndex        =   7
         Top             =   1040
         Width           =   1455
      End
      Begin VB.TextBox txtCasNo2 
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
         Left            =   3960
         TabIndex        =   9
         Top             =   1290
         Width           =   1455
      End
      Begin VB.TextBox txtCasNo1 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   1290
         Width           =   1455
      End
      Begin VB.TextBox TxtBatch1 
         Enabled         =   0   'False
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
         Left            =   7440
         TabIndex        =   14
         Top             =   1040
         Width           =   1455
      End
      Begin VB.TextBox TxtBatch2 
         Enabled         =   0   'False
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
         Left            =   9600
         TabIndex        =   15
         Top             =   1040
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Left            =   3480
         TabIndex        =   56
         Top             =   1050
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No. from"
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
         Left            =   240
         TabIndex        =   55
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Bordx Date from"
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
         Index           =   2
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
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
         Index           =   2
         Left            =   3480
         TabIndex        =   53
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label3 
         Caption         =   "Inv Reg Period"
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
         Left            =   6000
         TabIndex        =   52
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Batch No. from"
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
         Left            =   6000
         TabIndex        =   37
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label Label17 
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
         Left            =   9120
         TabIndex        =   36
         Top             =   1080
         Width           =   375
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
         Left            =   11160
         TabIndex        =   35
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label9 
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
         Left            =   3480
         TabIndex        =   34
         Top             =   820
         Width           =   375
      End
      Begin VB.Label Label7 
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
         Index           =   0
         Left            =   3480
         TabIndex        =   33
         Top             =   520
         Width           =   345
      End
      Begin VB.Label Label6 
         Caption         =   "Post Date from"
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
         Left            =   240
         TabIndex        =   32
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Post Status"
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
         Left            =   6000
         TabIndex        =   31
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Invoice Date from"
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
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   520
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Group Company"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Case Number from"
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
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   1300
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Index           =   0
         Left            =   3480
         TabIndex        =   27
         Top             =   1300
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstMBMInvList 
      Height          =   4455
      Left            =   60
      TabIndex        =   49
      Top             =   2280
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   7853
      View            =   3
      MultiSelect     =   -1  'True
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
         Text            =   "Chk"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "J/T"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "J/S"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "DB"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans.Ref."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AccPeriod"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Trans.Date"
         Object.Width           =   2469
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
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "D/C"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Clinics"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PLN Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "YYMM"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Grp Company"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Bordx Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "MemNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "BatchNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "PostIndex"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lstCDList 
      Height          =   3375
      Left            =   60
      TabIndex        =   50
      Top             =   3360
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   5948
      View            =   3
      MultiSelect     =   -1  'True
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
         Text            =   "Chk"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "J/T"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "J/S"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "DB"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans.Ref."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AccPeriod"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Trans.Date"
         Object.Width           =   2469
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
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "D/C"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Clinics"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "InsuredType"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Plan"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Grp Company"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Bordx Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Mem No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "ClinicCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "BatchNo"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Member Invoice Posting"
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
      TabIndex        =   51
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmPostingInvoiceMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Boolean
Dim PostingType As String
Dim PostingStatus As Boolean
Private m_blnDirection(1 To 22) As Boolean

Private Sub ComboListing()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String
    CboGrpCom.Clear
    tmpSelField = "SELECT AccMonth as tmpCode, AccYear as tmpName "
    tmpTable = "AccPeriod"
    tmpSelCriteria = "0=0"
    Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, cboAccPeriod, medixmasterconnMaint, "SearchMaintTableRec")
    cboPostStatus.AddItem "0 - No"
    cboPostStatus.AddItem "1 - Yes"
    cboPostStatus.ListIndex = 0
End Sub

Private Sub CboGrpCom_Change()
    If CboGrpCom.Text = "" Then
        If InStr(CboGrpCom.Text, "null") Then
           CboGrpCom.Enabled = False
        End If
    End If
End Sub

Private Sub CboGrpCom_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(CboGrpCom, KeyAscii)
End Sub

Private Sub cboInvReg_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(cboInvReg, KeyAscii)
End Sub
Private Sub cboPostStatus_KeyPress(KeyAscii As Integer)
    Call CboKeyPress(cboPostStatus, KeyAscii)
End Sub

Private Sub cboPostStatus_LostFocus()
    TxtBatch1.Enabled = False
    TxtBatch2.Enabled = False
    txtPostDate1.Enabled = False
    txtPostDate2.Enabled = False
    If Mid(cboPostStatus, 1, 1) = "1" Then
        TxtBatch1.Enabled = True
        TxtBatch2.Enabled = True
        txtPostDate1.Enabled = True
        txtPostDate2.Enabled = True
    End If
End Sub

Private Sub CheckAll_Click()
Dim lstCnt As Integer
    lstCnt = lstMBMInvList.ListItems.Count
    If LblPost = "Yes" Then
        CheckAll.Value = 0
        Exit Sub
    Else
        If lstCnt = 0 Then
            CheckAll.Value = 0
        Else
            For lstCnt = 1 To lstCnt
                If lstMBMInvList.ListItems(lstCnt).Checked = False Then
                    lstMBMInvList.ListItems(lstCnt).Checked = True
                    lstMBMInvList.ListItems(lstCnt).Text = "Post"
                ElseIf lstMBMInvList.ListItems(lstCnt).Checked = True Then
                    lstMBMInvList.ListItems(lstCnt).Checked = False
                    lstMBMInvList.ListItems(lstCnt).Text = ""
                End If
            Next lstCnt
        End If
    End If
End Sub

Private Sub cmdBack_Click()
    lstCDList.Visible = False
    lstMBMInvList.Visible = True
    CmdPost.Enabled = True
End Sub

Private Sub cmdchkAll_Click()
    Call SetCheckAllItems(True)
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    Call ClearInfo
    cboPostStatus.ListIndex = 0
    'CboBordx.ListIndex = 1
    FramePostInfo.Enabled = False
    CmdChkAll.Enabled = False
    TxtBatch1.Enabled = False
    TxtBatch2.Enabled = False
    txtPostDate1.Enabled = False
    txtPostDate2.Enabled = False
    lstMBMInvList.ListItems.Clear
    lstCDList.ListItems.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPost_Click()
Const METHOD_NAME = "CmdPost_Click"
On Error GoTo debugs

Dim AA As Long
Dim StrInvSql As String
Dim tmpClinic As String
Dim tmpCAsNo As String
Dim UpdAsk
Dim BatchNo As String
Dim CompTrans As Boolean
Dim ItemChecked As Boolean
Dim UpdateSuc As Boolean
    
    ItemChecked = False
    CompTrans = True
    UpdateSuc = False
    For AA = 1 To lstMBMInvList.ListItems.Count
        If lstMBMInvList.ListItems(AA).Checked = True Then
            ItemChecked = True
            Exit For
        End If
    Next AA
    If ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
        Exit Sub
    End If
    If LblPost = "Yes" Then
        MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    If txtBordxDateFr = "__/__/____" Or txtBordxDateTo = "__/__/____" Then
        MsgBox "Please Enter Bordx Date! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtBordxDateFr.SetFocus
        Exit Sub
    End If
    
    If Trim(cboAccPeriod.Text) = "" Then
        MsgBox "Please Keyin for Accounting Period ! ", vbOKOnly + vbCritical, DialogBoxTitle
        cboAccPeriod.SetFocus
        Exit Sub
    End If
    UpdAsk = MsgBox("Do you want to Post record(s) ?", vbYesNo + vbExclamation)
    If UpdAsk = vbYes Then
        Screen.MousePointer = vbHourglass
        On Err GoTo ErrorSave
        
        MediConnCISWS.beginTrans
        CompTrans = False
        BatchNo = GenSUNPostingBatchNo
        For AA = 1 To lstMBMInvList.ListItems.Count
            tmpCAsNo = Trim(lstMBMInvList.ListItems(AA).SubItems(4))
            tmpClinic = lstMBMInvList.ListItems(AA).SubItems(18)
            If lstMBMInvList.ListItems(AA).Checked = True And Trim(lstMBMInvList.ListItems(AA).SubItems(19)) = "Post" Then
                StrInvSql = ""
                StrInvSql = "Insert Into SUNPosting(InvoiceNo, TransRef, InvMBMStatus," & _
                            "PostStatus, PostType, PostDate, PostBy, " & _
                            "AccMonth, AccYear, BatchNo) " & _
                            "Values('" & tmpClinic & "','" & tmpCAsNo & "','1'" & _
                            ",'1','" & PostingType & "','" & Format(Date, "yyyy/mm/dd") & _
                            "','" & Logon & "','" & Mid(cboAccPeriod, 1, 2) & _
                            "','" & Mid(cboAccPeriod, 4, 4) & "','" & BatchNo & "')"
                MediConnCISWS.Execute StrInvSql
            End If
        Next AA
        MediConnCISWS.commitTrans
        CompTrans = True
        UpdateSuc = True
        lblBatchNo = BatchNo
        
        StrInvSql = ""
        StrInvSql = " and BatchNo = '" & Trim(lblBatchNo) & "'"
        Call MBMInvListing(StrInvSql, "G")
        PostingStatus = True
        Call CmdPrintFile_Click
        
        MsgBox "Record(s) Posted Sucessfully... ! ", vbOKOnly
        cboPostStatus.ListIndex = 1
        Screen.MousePointer = vbDefault
        End If
        
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            MediConnCISWS.RollbackTrans
            MsgBox Err.Description
        End If
        If Err.Description <> "" Then
            Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
        End If
        
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub CmdPrintFile_Click()
Dim SVRPath As String
Dim Filename As String
    'Call ExportToExcelAnalysis1(lstMBMInvList)
    
    Screen.MousePointer = vbHourglass
        
    If lstMBMInvList.ListItems.Count <> 0 Or lstCDList.ListItems.Count <> 0 Then
        If PostingStatus = False Then
            If lstMBMInvList.Visible = True Then
                Call ExportListViewtoExcel(lstMBMInvList)
            Else
                Call ExportListViewtoExcel(lstCDList)
            End If
        Else
                       '\\Mars\accounts\Account Journal Import\JI Not Posted
            SVRPath = "\\Mars\accounts\Account Journal Import\JIS CLN Not Posted\"
            Filename = lblBatchNo & " RC-MBMInv " & Format(txtBordxDateFr, "yyyymmdd") & "-" & Mid(txtBordxDateTo, 1, 2)
            Call ExporttoExcelPosting(lstMBMInvList, SVRPath, Filename)
            PostingStatus = False
        End If
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
    lstMBMInvList.ListItems.Clear
    lstMBMInvList.Visible = True
    lstCDList.Visible = False
    Call SearchRecord("G")
End Sub

Private Sub cmdSearchCredit_Click()
    lstMBMInvList.Visible = False
    lstCDList.Visible = True
    Call SearchRecord("C")
End Sub

Private Sub cmdSearchDebit_Click()
    lstMBMInvList.Visible = False
    lstCDList.Visible = True
    Call SearchRecord("D")
End Sub

Private Sub CmdStop_Click()
    intExit = True
    CmdExit.Enabled = True
End Sub
Private Sub Form_Load()
Call ActionLogUpdate(12, "frmPostingInvoiceMember")
    intExit = False
    PostingType = "InvMBM"
    Call ComboListing
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

Private Sub lstMBMInvList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "lstMBMInvList_ItemCheck"
On Error GoTo debugs
Dim CurrChk As Integer
Dim ModResult As Integer
Dim NextChk1 As Integer
Dim NextChk2 As Integer
Dim lstRec As ListItem
Dim PrevText As String
    NextChk2 = 0
    If LblPost = "Yes" Then
        MsgBox "Record(s) has been posted ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Item.Checked = False
    Else
        CurrChk = Item.Index
        If Item.ListSubItems(10) = "D" Then
            NextChk1 = CurrChk + 1
            If CurrChk + 2 <= lstMBMInvList.ListItems.Count Then
               If lstMBMInvList.ListItems(CurrChk + 2).SubItems(10) = "C" Then
                    NextChk2 = CurrChk + 2
               End If
            End If
        Else
            NextChk1 = CurrChk - 1
            If lstMBMInvList.ListItems(NextChk1).SubItems(10) = "C" Then
               NextChk2 = CurrChk - 2
            Else
                If CurrChk + 1 <= lstMBMInvList.ListItems.Count Then
                    If lstMBMInvList.ListItems(CurrChk + 1).SubItems(10) = "C" Then
                        NextChk2 = CurrChk + 1
                    End If
                End If
            End If
        End If
        PrevText = Item.Text
        If Item.Text = "" Then
            Item.Text = "Post"
            Item.Checked = True
            lstMBMInvList.ListItems(NextChk1).Text = "Post"
            lstMBMInvList.ListItems(NextChk1).Checked = True
            If NextChk2 > 0 Then
                lstMBMInvList.ListItems(NextChk2).Text = "Post"
                lstMBMInvList.ListItems(NextChk2).Checked = True
            End If
        End If
        If PrevText = "Post" Then
            Item.Text = ""
            Item.Checked = False
            lstMBMInvList.ListItems(NextChk1).Text = ""
            lstMBMInvList.ListItems(NextChk1).Checked = False
            If NextChk2 > 0 Then
                lstMBMInvList.ListItems(NextChk2).Text = ""
                lstMBMInvList.ListItems(NextChk2).Checked = False
            End If
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtBordxDateFr_LostFocus()
    If IsDate(txtBordxDateFr) Then
    Else
        If txtBordxDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateFr.SetFocus
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

Private Sub txtCasNo1_LostFocus()
    txtCasNo1 = ChkStr(txtCasNo1)
End Sub

Private Sub txtCasNo2_LostFocus()
    txtCasNo2 = ChkStr(txtCasNo2)
End Sub

Private Sub txtInvDate1_LostFocus()
    If IsDate(txtInvDate1) Then
    Else
        If txtInvDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtInvDate2_LostFocus()
    If IsDate(txtInvDate2) Then
    Else
        If txtInvDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtPostDate1_LostFocus()
    If IsDate(txtPostDate1) Then
    Else
        If txtPostDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPostDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtPostDate2_LostFocus()
    If IsDate(txtPostDate2) Then
    Else
        If txtPostDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPostDate2.SetFocus
        End If
    End If
End Sub

Private Sub MBMInvListing(strAppend As String, tmpCDType As String)
Const METHOD_NAME = "MBMInvListing"
On Error GoTo debugs

Dim MBMInv As New ADODB.Recordset
Dim JournalType, AcCode, CreditType As String
Dim AA, CurrRec As Integer
Dim TotDbAmt, TotCrAmt As Double
Dim cmd As New ADODB.Command
Dim App1 As New ADODB.Parameter
Dim App2 As New ADODB.Parameter
Dim RecLoop As Integer
Dim tmpDesc As String
Dim lstRec As ListItem
Dim tmpMBMNo As String
Dim tmpPos As Integer
Dim Last4MBMNo As String
Dim tmpPostageFee As Double
Dim SuppRec As New ADODB.Recordset
Dim StrSql As String
Dim CrCnt As Integer
Dim tmpMCOProcFee As Double
Dim tmpDispAmt As Double
Dim tmpColPayAmt As Double
Dim tmpActualAmt As Double
Dim RegMonth As String

    CurrRec = 0
    LblPost = "No"
    If tmpCDType = "G" Then
         lstMBMInvList.ListItems.Clear
    Else
         lstCDList.ListItems.Clear
    End If
    'VIEW_SUN_Posting_InvMBM
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "SUN_CIS_Posting"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set App1 = cmd.CreateParameter("RptType", adVarChar, adParamInput, 20, PostingType)
    cmd.Parameters.Append App1
    Set App2 = cmd.CreateParameter("StrCriterial", adVarChar, adParamInput, 1000, strAppend)
    cmd.Parameters.Append App2
    MBMInv.CursorLocation = adUseClient
    MBMInv.CursorType = adOpenForwardOnly
    MBMInv.CacheSize = 100
    Set MBMInv = cmd.Execute
    Do While Not MBMInv.EOF
       With MBMInv
            CurrRec = CurrRec + 1
            JournalType = "RC"
            RecLoop = 1
            If tmpCDType = "G" Then RecLoop = 2
            tmpColPayAmt = IIf(IsNumeric(!CLMCoPayColl), !CLMCoPayColl, 0)
            tmpPostageFee = IIf(IsNumeric(!CLMPostageFee), !CLMPostageFee, 0)
            tmpActualAmt = IIf(IsNumeric(!CLMTotalActual), !CLMTotalActual, 0)
            If tmpPostageFee > 0 And tmpCDType <> "D" Then
                RecLoop = RecLoop + 1
            End If
            tmpPos = InStr(1, !MBMNumber, "*")
            If tmpPos = 0 Then
                Last4MBMNo = Right(!MBMNumber, 4)
            Else
                tmpMBMNo = Mid(!MBMNumber, 1, tmpPos - 1)
                Last4MBMNo = Right(tmpMBMNo, 4)
            End If
            tmpDesc = ""
            If Trim(!coverid) = "00" Then
                tmpDesc = !MBMName
            Else
                StrSql = "SELECT MBMCName from MBMCoveredPersons " & _
                         " Where MBMCNumber='" & Trim(!MBMNumber) & "' AND MBMCCoverID = '" & Trim(!coverid) & "'"
                SuppRec.CursorLocation = adUseClient
                SuppRec.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                Do While Not SuppRec.EOF
                    tmpDesc = IIf(Len(SuppRec!MBMCName), Trim(SuppRec!MBMCName), "")
                    SuppRec.MoveNext
                Loop
                SuppRec.Close
                Set SuppRec = Nothing
            End If
            If MBMInv!InvMBMStatus = "1" Then
                LblPost = "Yes"
            End If
            
            AcCode = "6000"
            CreditType = "D"
            tmpDispAmt = tmpActualAmt
            If tmpCDType = "C" Then
                AcCode = "82" & !GRPCoCode
                CreditType = "C"
            End If
            
            CrCnt = 0
            For AA = 1 To RecLoop
                If AA > 1 Then
                    CrCnt = CrCnt + 1
                    If CrCnt = 1 Then
                        AcCode = "82" & !GRPCoCode
                        tmpDispAmt = tmpActualAmt - tmpPostageFee
                    ElseIf CrCnt = 2 Then
                        AcCode = "4040"
                        tmpDispAmt = tmpPostageFee
                    End If
                    CreditType = "C"
                    If tmpCDType = "G" Then
                        If AA = RecLoop Then
                            lstRec.SubItems(19) = "Post"
                        End If
                    End If
                End If
                If tmpCDType = "G" Then
                    Set lstRec = lstMBMInvList.ListItems.Add(, , IIf(MBMInv!InvMBMStatus = 1, "Post", ""))
                Else
                    Set lstRec = lstCDList.ListItems.Add(, , IIf(MBMInv!InvMBMStatus = 1, "Post", ""))
                End If
                
                lstRec.SubItems(1) = JournalType
                'lstRec.SubItems(2) J/S
                lstRec.SubItems(3) = "UHC"
                lstRec.SubItems(4) = Trim(!CasNumber)
                If Trim(!AccMonth & !AccYear) <> "" Then
                    lstRec.SubItems(5) = !AccMonth & "/" & !AccYear
                End If
                lstRec.SubItems(6) = IIf(IsDate(!InvoiceDate), Format(!InvoiceDate, "DD/MM/YYYY"), "")
                lstRec.SubItems(7) = Mid(Trim(tmpDesc), 1, 25)
                lstRec.SubItems(8) = IIf(Len(AcCode), AcCode, "")
                lstRec.SubItems(9) = Format(tmpDispAmt, "#,##0.00")
                lstRec.SubItems(10) = Trim(CreditType)
                lstRec.SubItems(11) = "1R"
                'lstRec.SubItems(12) Plan
                lstRec.SubItems(13) = IIf(Len(!PAYCode), Trim(!PAYCode), "")
                
                If Len(!BordxDate) Then
                    lstRec.SubItems(14) = Mid(!BordxDate, 9, 2) & Mid(!BordxDate, 4, 2)
                Else
                    lstRec.SubItems(14) = ""
                End If
                
                lstRec.SubItems(15) = IIf(Len(!GRPCoCode), Trim(!GRPCoCode), "")
                'lstRec.SubItems(16) = IIf(Len(!INVBordxRefNo), Right(!INVBordxRefNo, 5), "")
                lstRec.SubItems(16) = IIf(Len(!INVBordxRefNo), Trim(!INVBordxRefNo), "")
                lstRec.SubItems(17) = Trim(Last4MBMNo)
                lstRec.SubItems(18) = IIf(Len(Trim(!BatchNo)), Trim(!BatchNo), "")
                If CreditType = "C" Then
                    TotCrAmt = TotCrAmt + tmpDispAmt
                Else
                    TotDbAmt = TotDbAmt + tmpDispAmt
                End If
            Next AA
            LblDbAmt = Format(TotDbAmt, "#,##0.00")
            LblCrAmt = Format(TotCrAmt, "#,##0.00")
            lblCurrent = CurrRec
        End With

        MBMInv.MoveNext
        DoEvents
        If intExit = True Then
           Exit Do
        End If
    Loop
    intExit = 0
    
    MBMInv.Close
    Set MBMInv = Nothing
    If CurrRec = 0 Then
        MsgBox "No record selected ! ", vbOKOnly + vbCritical
    ElseIf LblPost = "No" Then
        FramePostInfo.Enabled = True
        CmdChkAll.Enabled = True
    End If
    cmdSearchCredit.Enabled = True
    cmdSearchDebit.Enabled = True
    CmdPost.Enabled = True
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub ClearInfo()
    lblCurrent = ""
    LblDbAmt = ""
    LblCrAmt = ""
    lblBatchNo = ""
    'cboAccPeriod.Text = ""
    FramePostInfo.Enabled = False
    CmdChkAll.Enabled = False
    CheckAll.Value = 0
    cmdSearchCredit.Enabled = False
    cmdSearchDebit.Enabled = False
    CmdPost.Enabled = False
End Sub

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstMBMInvList.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstMBMInvList.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub
Private Sub ChkGrp_Click()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String
    CboGrpCom.Clear
    If ChkGrp.Value = 1 Then
            tmpSelField = "SELECT GRPCoCode as tmpCode, GRPCoName as tmpName "
            tmpTable = "GRPCompany"
            tmpSelCriteria = ""
            Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, CboGrpCom, MediConnCISDB, "SearchCISDBTableRec")
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
        ElseIf tmpTableName = "AccPeriod" Then
            tmpComboBox.AddItem !TmpCode & "/" & !tmpName
            cboInvReg.AddItem !TmpCode & "/" & !tmpName
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

Private Function SearchRecord(tmpStatus As String)
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim StrSql As String
Dim PostSql As String

    Call ClearInfo
    StrSql = ""
    PostSql = ""
    If Trim(CboGrpCom) <> "" Then
        StrSql = StrSql & " And MBMGroupCompany='" & Trim(CboGrpCom) & "'"
    End If
    If IsDate(txtInvDate1) Then
        StrSql = StrSql & " And InvoiceDate >='" & Format(txtInvDate1, "yyyy/mm/dd") & "'"
    End If
    If IsDate(txtInvDate2) Then
        StrSql = StrSql & " And InvoiceDate <='" & Format(txtInvDate2, "yyyy/mm/dd") & "'"
    End If
    If Trim(txtCasNo1) <> "" Then
        StrSql = StrSql & " And CasNumber >='" & Trim(txtCasNo1) & "'"
    End If
    If Trim(txtCasNo2) <> "" Then
        StrSql = StrSql & " And CasNumber <='" & Trim(txtCasNo2) & "'"
    End If
    If Mid(cboPostStatus, 1, 1) = "0" Then
        PostSql = " And (InvMBMStatus IS NULL OR InvMBMStatus = '0')"
        PostSql = PostSql & " And (PostStatus IS NULL OR PostStatus = '0')"
    ElseIf Mid(cboPostStatus, 1, 1) = "1" Then
        PostSql = " And InvMBMStatus ='1' And PostStatus ='1'"
    End If
    If IsDate(txtPostDate1) Then
        StrSql = StrSql & " And PostDate >='" & Format(txtPostDate1, "yyyy/mm/dd") & "'"
    End If
    If IsDate(txtPostDate2) Then
        StrSql = StrSql & " And PostDate <='" & Format(txtPostDate2, "yyyy/mm/dd") & "'"
    End If
    If Trim(TxtBatch1) <> "" Then
        StrSql = StrSql & " And BatchNo >='" & Trim(TxtBatch1) & "'"
    End If
    If Trim(TxtBatch2) <> "" Then
        StrSql = StrSql & " And BatchNo <='" & Trim(TxtBatch2) & "'"
    End If

    If IsDate(txtBordxDateFr) Then
        StrSql = StrSql & " And BordxDate >='" & Format(txtBordxDateFr, "yyyy/mm/dd") & "'"
    End If
    If IsDate(txtBordxDateTo) Then
        StrSql = StrSql & " And BordxDate <='" & Format(txtBordxDateTo, "yyyy/mm/dd") & "'"
    End If

    If Len(txtBordxNoFr) Then
        StrSql = StrSql & " And BordxRefNo >='" & Trim(txtBordxNoFr) & "'"
    End If
    If Len(txtBordxNoTo) Then
        StrSql = StrSql & " And BordxRefNo <='" & Trim(txtBordxNoTo) & "'"
    End If

    If Len(cboInvReg) Then
        StrSql = StrSql & " And RegMonth ='" & Mid(cboInvReg, 1, 2) & "'"
        StrSql = StrSql & " And RegYr ='" & Mid(cboInvReg, 4, 4) & "'"
    End If
    
    Screen.MousePointer = vbHourglass
        Call MBMInvListing(StrSql & PostSql, tmpStatus)
    Screen.MousePointer = vbDefault

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Function

Private Sub CboKeyPress(tmpCbo As ComboBox, KeyAscii As Integer)
Const METHOD_NAME = "CboKeyPress"
On Error GoTo debugs

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
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub


