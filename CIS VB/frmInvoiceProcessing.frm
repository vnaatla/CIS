VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInvoiceProcessing 
   Caption         =   "Invoice Processing"
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
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   2085
      Left            =   240
      TabIndex        =   136
      Top             =   5400
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3678
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   46
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "WS Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Inv No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Inv Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Clinic"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Claim Cat"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "INV Amount"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Pay To"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Payee"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Claim Type"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PreparedBy"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "GP Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "SP Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Medicine"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Procedures"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Lab"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "X-Ray"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "MRI"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "Dental"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "Chinese Med"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Text            =   "Traditional Med"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "Physio Theraphy"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "Accident"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "PRV Checkup"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "Vision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Text            =   "SSTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Text            =   "Others1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Text            =   "Others2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Text            =   "NC GP Fees"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Text            =   "NC Medicine"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   33
         Text            =   "NC Medical Rpt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   34
         Text            =   "NC Others"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Claimability"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "Pending"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "Reason1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "Reason2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "Reason3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "MC"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "RepCode1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "RepCode2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "RepCode3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "Bordxed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "WS Status"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame FrameInvoice 
      Height          =   1575
      Left            =   6240
      TabIndex        =   128
      Top             =   1040
      Width           =   5535
      Begin MSMask.MaskEdBox txtInvoiceDate 
         Height          =   315
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox CboPAYto 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Text            =   "C - Clinic"
         Top             =   520
         Width           =   1695
      End
      Begin VB.TextBox TxtInvPayType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3960
         TabIndex        =   129
         Text            =   "Cashless"
         Top             =   520
         Width           =   1455
      End
      Begin VB.ComboBox Cboclinic 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   810
         Width           =   4335
      End
      Begin VB.ComboBox CboPayee 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   1100
         Width           =   4335
      End
      Begin VB.Label Label11 
         Caption         =   "Cash Type"
         Height          =   255
         Left            =   2920
         TabIndex        =   135
         Top             =   560
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Payee"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Pay To"
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   2920
         TabIndex        =   132
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Clinic"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame FrameCase 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5640
      TabIndex        =   113
      Top             =   400
      Width           =   6135
      Begin VB.ComboBox cboPending 
         Height          =   315
         Left            =   5160
         TabIndex        =   47
         Text            =   "N - No"
         Top             =   200
         Width           =   855
      End
      Begin VB.ComboBox cboStatus 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   45
         Text            =   "O - Open"
         Top             =   200
         Width           =   975
      End
      Begin VB.ComboBox cboClaimability 
         Height          =   315
         Left            =   3120
         TabIndex        =   46
         Text            =   "A - Accepted"
         Top             =   200
         Width           =   1215
      End
      Begin VB.Label Label49 
         Caption         =   "Pending"
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
         Left            =   4440
         TabIndex        =   116
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label42 
         Caption         =   "Case Status"
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
         TabIndex        =   115
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label44 
         Caption         =   "Claimability"
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
         Left            =   2160
         TabIndex        =   114
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Invoice Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   80
      TabIndex        =   67
      Top             =   1040
      Width           =   6015
      Begin VB.CommandButton cmdShowINv 
         Appearance      =   0  'Flat
         Caption         =   ">"
         Height          =   270
         Left            =   5600
         TabIndex        =   6
         Top             =   550
         Width           =   285
      End
      Begin VB.TextBox txtMBMName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   255
         Width           =   4695
      End
      Begin VB.ComboBox CboPatientID 
         Height          =   315
         ItemData        =   "frmInvoiceProcessing.frx":0000
         Left            =   1200
         List            =   "frmInvoiceProcessing.frx":0002
         TabIndex        =   4
         Top             =   510
         Width           =   2175
      End
      Begin MSMask.MaskEdBox txtDateVisit 
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtPatientDOB 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   111
         Top             =   800
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Date Visit"
         Height          =   255
         Left            =   3480
         TabIndex        =   126
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label46 
         Caption         =   "Patient DOB"
         Height          =   255
         Left            =   240
         TabIndex        =   109
         Top             =   820
         Width           =   975
      End
      Begin VB.Label Label45 
         Caption         =   "MBM Name"
         Height          =   255
         Left            =   240
         TabIndex        =   108
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Patient"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   520
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Member No"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FrameCLMCat 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Claim Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   80
      TabIndex        =   66
      Top             =   2280
      Width           =   6015
      Begin VB.TextBox txtMC 
         Height          =   285
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox CboClmCat 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Text            =   "GP - GP Clinic Consultation"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label16 
         Caption         =   "Days"
         Height          =   255
         Left            =   5400
         TabIndex        =   153
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Medical Certificate (MC)"
         Height          =   255
         Left            =   2880
         TabIndex        =   152
         Top             =   260
         Width           =   1815
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   120
      TabIndex        =   57
      Top             =   7560
      Width           =   11655
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   3360
         TabIndex        =   51
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2520
         TabIndex        =   50
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdTotal 
         Caption         =   "&Total"
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   48
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10560
         TabIndex        =   54
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9720
         TabIndex        =   53
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdClaerInv 
         Caption         =   "Clear &Inv"
         Height          =   330
         Left            =   8880
         TabIndex        =   52
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Enabled         =   0   'False
         Height          =   330
         Left            =   960
         TabIndex        =   49
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   330
         Left            =   960
         TabIndex        =   117
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1680
         TabIndex        =   127
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label54 
         Caption         =   "of"
         Height          =   255
         Left            =   5760
         TabIndex        =   65
         Top             =   200
         Width           =   255
      End
      Begin VB.Label lblLastVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   64
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label58 
         Caption         =   "Bordx?"
         Height          =   255
         Left            =   6600
         TabIndex        =   63
         Top             =   200
         Width           =   615
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   62
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label50 
         Caption         =   "Version"
         Height          =   255
         Left            =   4680
         TabIndex        =   61
         Top             =   200
         Width           =   615
      End
      Begin VB.Label lblBordx 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   7320
         TabIndex        =   60
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label63 
         Caption         =   "WS?"
         Height          =   255
         Left            =   7800
         TabIndex        =   59
         Top             =   200
         Width           =   495
      End
      Begin VB.Label lblWs 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   8280
         TabIndex        =   58
         Top             =   200
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   80
      TabIndex        =   55
      Top             =   400
      Width           =   5415
      Begin VB.CheckBox ChkKey 
         Height          =   255
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4620
         TabIndex        =   72
         Top             =   200
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Height          =   330
         Left            =   3900
         TabIndex        =   2
         Top             =   200
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtKeyNo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   200
         Width           =   2295
      End
      Begin VB.Label lblkey 
         Caption         =   "Member No."
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4545
      Left            =   80
      TabIndex        =   73
      Top             =   3000
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8017
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Invoice Entry"
      TabPicture(0)   =   "frmInvoiceProcessing.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtINVAmount"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameCharges"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameReason"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Not Cover"
      TabPicture(1)   =   "frmInvoiceProcessing.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameNCover"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "General Info"
      TabPicture(2)   =   "frmInvoiceProcessing.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameMBMInfo"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Case Info"
      TabPicture(3)   =   "frmInvoiceProcessing.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameCASeInfo"
      Tab(3).ControlCount=   1
      Begin VB.Frame FrameReason 
         Caption         =   "Reason For Seeking Treatment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1260
         Left            =   7680
         TabIndex        =   155
         Top             =   480
         Width           =   3945
         Begin VB.ComboBox CboReason1 
            Height          =   315
            Left            =   260
            TabIndex        =   34
            Tag             =   "0"
            Top             =   240
            Width           =   3600
         End
         Begin VB.ComboBox CboReason2 
            Height          =   315
            Left            =   260
            TabIndex        =   35
            Top             =   520
            Width           =   3600
         End
         Begin VB.ComboBox CboReason3 
            Height          =   315
            Left            =   260
            TabIndex        =   36
            Top             =   800
            Width           =   3600
         End
         Begin VB.Label Label52 
            Caption         =   "3"
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
            TabIndex        =   158
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label51 
            Caption         =   "2"
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
            TabIndex        =   157
            Top             =   525
            Width           =   255
         End
         Begin VB.Label Label12 
            Caption         =   "1"
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
            TabIndex        =   156
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.Frame FrameCASeInfo 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   142
         Top             =   360
         Width           =   5265
         Begin VB.Frame FrameRepudiate 
            Caption         =   "Repudiation Reasons"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   143
            Top             =   800
            Width           =   4995
            Begin VB.ComboBox CboRepudiate1 
               Height          =   315
               Left            =   480
               TabIndex        =   42
               Top             =   240
               Width           =   4455
            End
            Begin VB.ComboBox CboRepudiate2 
               Height          =   315
               Left            =   480
               TabIndex        =   43
               Top             =   480
               Width           =   4455
            End
            Begin VB.ComboBox CboRepudiate3 
               Height          =   315
               Left            =   480
               TabIndex        =   44
               Top             =   720
               Width           =   4455
            End
            Begin VB.Label Label48 
               Caption         =   "3"
               Height          =   255
               Left            =   240
               TabIndex        =   146
               Top             =   780
               Width           =   255
            End
            Begin VB.Label Label47 
               Caption         =   "2"
               Height          =   180
               Left            =   240
               TabIndex        =   145
               Top             =   525
               Width           =   255
            End
            Begin VB.Label Label43 
               Caption         =   "1"
               Height          =   255
               Left            =   240
               TabIndex        =   144
               Top             =   260
               Width           =   255
            End
         End
         Begin VB.TextBox TxtRemarks 
            Height          =   570
            Left            =   840
            MaxLength       =   100
            TabIndex        =   41
            Top             =   210
            Width           =   4215
         End
         Begin VB.Label Label41 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   195
            Width           =   855
         End
      End
      Begin VB.Frame FrameMBMInfo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1875
         Left            =   -74880
         TabIndex        =   92
         Top             =   360
         Width           =   6015
         Begin VB.TextBox txtTabMBMName 
            Height          =   285
            Left            =   1320
            TabIndex        =   93
            Top             =   240
            Width           =   4575
         End
         Begin VB.TextBox TxtPayor 
            Height          =   285
            Left            =   6840
            TabIndex        =   121
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTabMBMNo 
            Height          =   285
            Left            =   1320
            TabIndex        =   118
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox TxtPolNo 
            Height          =   285
            Left            =   1320
            TabIndex        =   99
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtdob 
            Height          =   285
            Left            =   4680
            TabIndex        =   110
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtPolStatus 
            Height          =   285
            Left            =   4680
            TabIndex        =   98
            Top             =   720
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtEffDate 
            Height          =   285
            Left            =   4680
            TabIndex        =   94
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtExpDate 
            Height          =   315
            Left            =   4680
            TabIndex        =   95
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtClmTpye 
            Height          =   285
            Left            =   4680
            TabIndex        =   97
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtPlanCode 
            Height          =   285
            Left            =   1320
            TabIndex        =   125
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox TxtInsType 
            Height          =   285
            Left            =   1320
            TabIndex        =   124
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtNatureCLM 
            Height          =   285
            Left            =   1320
            TabIndex        =   96
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label64 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   1020
            Width           =   975
         End
         Begin VB.Label Label62 
            Caption         =   "Insured Type"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label61 
            Caption         =   "Payor"
            Height          =   255
            Left            =   6120
            TabIndex        =   120
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label60 
            Caption         =   "Member No"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label40 
            Caption         =   "ClmType"
            Height          =   255
            Left            =   3600
            TabIndex        =   107
            Top             =   1520
            Width           =   855
         End
         Begin VB.Label Label39 
            Caption         =   "Nature of CLM"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   1520
            Width           =   1095
         End
         Begin VB.Label Label38 
            Caption         =   "Pol. Status"
            Height          =   255
            Left            =   3600
            TabIndex        =   105
            Top             =   800
            Width           =   855
         End
         Begin VB.Label Label37 
            Caption         =   "Patient DOB"
            Height          =   255
            Left            =   3600
            TabIndex        =   104
            Top             =   560
            Width           =   975
         End
         Begin VB.Label Label36 
            Caption         =   "Exp Date"
            Height          =   255
            Left            =   3600
            TabIndex        =   103
            Top             =   1280
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   3600
            TabIndex        =   102
            Top             =   1040
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Policy No"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   765
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Member Name"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame FrameCharges 
         Caption         =   "Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1935
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   7455
         Begin VB.TextBox A1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1120
            MaxLength       =   8
            TabIndex        =   14
            Top             =   240
            Width           =   680
         End
         Begin VB.TextBox A2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1120
            MaxLength       =   8
            TabIndex        =   15
            Top             =   480
            Width           =   680
         End
         Begin VB.TextBox A3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1120
            MaxLength       =   8
            TabIndex        =   16
            Top             =   720
            Width           =   680
         End
         Begin VB.TextBox A4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1120
            MaxLength       =   8
            TabIndex        =   17
            Top             =   960
            Width           =   680
         End
         Begin VB.TextBox A5 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   18
            Top             =   240
            Width           =   680
         End
         Begin VB.TextBox A6 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   19
            Top             =   480
            Width           =   680
         End
         Begin VB.TextBox A7 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   20
            Top             =   720
            Width           =   680
         End
         Begin VB.TextBox A8 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   21
            Top             =   960
            Width           =   680
         End
         Begin VB.TextBox A9 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4800
            MaxLength       =   8
            TabIndex        =   22
            Top             =   240
            Width           =   680
         End
         Begin VB.TextBox A10 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4800
            MaxLength       =   8
            TabIndex        =   23
            Top             =   480
            Width           =   680
         End
         Begin VB.TextBox A11 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4800
            MaxLength       =   8
            TabIndex        =   24
            Top             =   720
            Width           =   680
         End
         Begin VB.TextBox A12 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4800
            MaxLength       =   8
            TabIndex        =   25
            Top             =   960
            Width           =   680
         End
         Begin VB.TextBox A13 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   26
            Top             =   240
            Width           =   680
         End
         Begin VB.TextBox A14 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   27
            Top             =   480
            Width           =   680
         End
         Begin VB.TextBox A15 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   28
            Top             =   720
            Width           =   680
         End
         Begin VB.TextBox A16 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   29
            Top             =   960
            Width           =   680
         End
         Begin VB.TextBox A17 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   31
            Top             =   1200
            Width           =   680
         End
         Begin VB.TextBox TxtIndex 
            Height          =   195
            Left            =   960
            TabIndex        =   148
            Top             =   960
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtOtherDesc1 
            Height          =   285
            Left            =   1120
            MaxLength       =   100
            TabIndex        =   30
            Top             =   1200
            Width           =   4350
         End
         Begin VB.TextBox txtOtherDesc2 
            Height          =   285
            Left            =   1120
            MaxLength       =   100
            TabIndex        =   32
            Top             =   1440
            Width           =   4350
         End
         Begin VB.TextBox A18 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   33
            Top             =   1440
            Width           =   680
         End
         Begin VB.Label Label65 
            Caption         =   "Amount"
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
            Left            =   5640
            TabIndex        =   151
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label53 
            Caption         =   "Amount"
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
            Left            =   5640
            TabIndex        =   150
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Others2 Desc"
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
            TabIndex        =   149
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "GP Cons Fee"
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
            TabIndex        =   91
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "SP Cons Fee"
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
            TabIndex        =   90
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "Medication"
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
            TabIndex        =   89
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Procedures"
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
            TabIndex        =   88
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "Lab"
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
            Left            =   1920
            TabIndex        =   87
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "X-Ray"
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
            Left            =   1920
            TabIndex        =   86
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label24 
            Caption         =   "MRI"
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
            Left            =   1920
            TabIndex        =   85
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "Dental"
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
            Left            =   1920
            TabIndex        =   84
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "Chinese Medicine"
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
            TabIndex        =   83
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Traditional Med"
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
            TabIndex        =   82
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "Physio Theraphy"
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
            TabIndex        =   81
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Accident"
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
            TabIndex        =   80
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "PRV Checkup"
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
            Left            =   5640
            TabIndex        =   79
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label31 
            Caption         =   "Vision"
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
            Left            =   5640
            TabIndex        =   78
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label32 
            Caption         =   "I&&D"
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
            Left            =   5640
            TabIndex        =   77
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label33 
            Caption         =   "T&&S / STO"
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
            Left            =   5640
            TabIndex        =   76
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label34 
            Caption         =   "Others1 Desc"
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
            TabIndex        =   75
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.Frame FrameNCover 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   137
         Top             =   480
         Width           =   3855
         Begin VB.TextBox TxtNCGPFee 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TxtNCMedicine 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   38
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TxtNCMedicalRPT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   39
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TxtNCOthers 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   40
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label55 
            Caption         =   "GP Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   141
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label56 
            Caption         =   "Medicine"
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   615
            Width           =   975
         End
         Begin VB.Label Label57 
            Caption         =   "Medical Report"
            Height          =   255
            Left            =   240
            TabIndex        =   139
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label Label59 
            Caption         =   "Others"
            Height          =   255
            Left            =   240
            TabIndex        =   138
            Top             =   1185
            Width           =   975
         End
      End
      Begin VB.Label txtINVAmount 
         Height          =   375
         Left            =   7680
         TabIndex        =   154
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Label LblExpDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9960
      TabIndex        =   112
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label35 
      BackColor       =   &H00404040&
      Caption         =   "  Invoice Processing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   400
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmInvoiceProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bExit As Boolean
Dim GenCaseNo As Boolean
Dim intIndex As String
Dim EditFlag As Boolean
Dim AddFlag As Boolean
Dim checking As Boolean
Dim HLTCode, INSCode, PLNCode As String
'Dim objword As Word.Application
'Dim objdoc As Word.Document
Dim ScanCode As String
Dim DOBArr(0 To 20)
Dim CASENumber, NewCaseNo, MembershipNo As String
Private Hospitals As String

Private Sub A1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A8_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A9_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A10_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A11_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A12_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A13_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A14_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A15_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A16_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A17_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A18_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub cboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClaimability.Text, cboClaimability.SelStart) + Chr(KeyAscii) + Mid(cboClaimability.Text, cboClaimability.SelStart + cboClaimability.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboClaimability.ListCount - 1
       
        If NewText = LCase(Left(cboClaimability.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboClaimability.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboClaimability.Text = ValidValue
            cboClaimability.SelStart = 0
            cboClaimability.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub cboClaimability_LostFocus()
    If Mid(cboClaimability, 1, 1) = "R" Then
        FrameRepudiate.Enabled = True
    End If
End Sub

Private Sub Cboclinic_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(Cboclinic.Text, Cboclinic.SelStart) + Chr(KeyAscii) + Mid(Cboclinic.Text, Cboclinic.SelStart + Cboclinic.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To Cboclinic.ListCount - 1
       
        If NewText = LCase(Left(Cboclinic.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = Cboclinic.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            Cboclinic.Text = ValidValue
            Cboclinic.SelStart = 0
            Cboclinic.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboClmCat_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClmCat.Text, CboClmCat.SelStart) + Chr(KeyAscii) + Mid(CboClmCat.Text, CboClmCat.SelStart + CboClmCat.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboClmCat.ListCount - 1
       
        If NewText = LCase(Left(CboClmCat.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboClmCat.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboClmCat.Text = ValidValue
            CboClmCat.SelStart = 0
            CboClmCat.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboPatientID_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPatientID.Text, CboPatientID.SelStart) + Chr(KeyAscii) + Mid(CboPatientID.Text, CboPatientID.SelStart + CboPatientID.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboPatientID.ListCount - 1
       
        If NewText = LCase(Left(CboPatientID.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboPatientID.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPatientID.Text = ValidValue
            CboPatientID.SelStart = 0
            CboPatientID.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboPAYto_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPAYto.Text, CboPAYto.SelStart) + Chr(KeyAscii) + Mid(CboPAYto.Text, CboPAYto.SelStart + CboPAYto.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboPAYto.ListCount - 1
       
        If NewText = LCase(Left(CboPAYto.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboPAYto.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPAYto.Text = ValidValue
            CboPAYto.SelStart = 0
            CboPAYto.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub cboPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPending.Text, cboPending.SelStart) + Chr(KeyAscii) + Mid(cboPending.Text, cboPending.SelStart + cboPending.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPending.ListCount - 1
       
        If NewText = LCase(Left(cboPending.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboPending.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboPending.Text = ValidValue
            cboPending.SelStart = 0
            cboPending.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub CmdCancel_Click()
    Call EntriesEnable(False)
    CmdAdd.Enabled = True
    CmdDelete.Visible = False
    CmdEdit.Visible = True
    CmdTotal.Enabled = True
End Sub

Private Sub CmdClaerInv_Click()
    Call ClearInvoice
End Sub

Private Sub cmdDelete_Click()

On Error GoTo errorRemove
    
    Dim ForLoop As Integer
    Dim beginTrans As Boolean
    Dim ListCount As Integer
    Dim strsql, strActual, strNot As String
    Dim RecordSaved
    Dim RemoveInvoice  As New ADODB.Recordset
    Dim RemoveActualInvoice  As New ADODB.Recordset
    Dim RemoveNotCover  As New ADODB.Recordset
    
    If lstInvoice.SelectedItem.SubItems(3) = False Then
        
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
          If RecordSaved = vbYes Then
              ListCount = lstInvoice.ListItems.count
              For ForLoop = 1 To ListCount
                  If lstInvoice.ListItems(ForLoop).Selected = True Then
                      lstInvoice.ListItems.Remove (ForLoop)
                      Exit For
                  End If
              Next ForLoop
              
            strsql = "delete  from ClaimInvoice where INVIndex =" & TxtIndex
            MediConnCISWS.beginTrans
            beginTrans = True
            RemoveInvoice.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
            strActual = "delete from ClaimInvoiceActual where INVIndex =" & TxtIndex
            RemoveActualInvoice.Open strActual, MediConnCISWS, adOpenKeyset, adLockOptimistic
            strNot = "delete from ClaimInvoiceNotCover where INVIndex =" & TxtIndex
            RemoveNotCover.Open strNot, MediConnCISWS, adOpenKeyset, adLockOptimistic
            
            MediConnCISWS.CommitTrans
          End If
          checking = False
          Call ClearInvoice
          checking = True
    Else
        MsgBox "Selected Invoice cannot be deleted!" & vbNewLine & "It has been used in Worksheet!", vbCritical
    End If
    Exit Sub
errorRemove:
    If beginTrans = True Then
        MediConnCISWS.RollbackTrans
    End If
    MsgBox Err.Description
End Sub

Private Sub cmdShowINv_Click()
    If IsDate(txtDateVisit) = False Or Mid(CboPatientID, 1, 2) = "" Then
            MsgBox "Invalid date format OR pls keyin Patient ID !!", vbCritical
    ElseIf Mid(CboPatientID, 1, 2) <> "" Then
            Call InvoiceINFO
    End If
    If lstInvoice.ListItems.count <= 0 Then
        CmdAdd.Enabled = True
    End If
    
End Sub

Private Sub CmdTotal_Click()
Dim AA, BB, ListCnt As Integer
    
    CmdEdit.Enabled = False
    CmdAdd.Enabled = True
    CmdDelete.Enabled = False
    
    ListCount = lstInvoice.ListItems.count
    A1 = CDbl(IIf(IsNumeric(A1), A1, 0))
    A2 = CDbl(IIf(IsNumeric(A2), A2, 0))
    A3 = CDbl(IIf(IsNumeric(A3), A3, 0))
    A4 = CDbl(IIf(IsNumeric(A4), A4, 0))
    A5 = CDbl(IIf(IsNumeric(A5), A5, 0))
    A6 = CDbl(IIf(IsNumeric(A6), A6, 0))
    A7 = CDbl(IIf(IsNumeric(A7), A7, 0))
    A8 = CDbl(IIf(IsNumeric(A8), A8, 0))
    A9 = CDbl(IIf(IsNumeric(A9), A9, 0))
    A10 = CDbl(IIf(IsNumeric(A10), A10, 0))
    A11 = CDbl(IIf(IsNumeric(A11), A11, 0))
    A12 = CDbl(IIf(IsNumeric(A12), A12, 0))
    A13 = CDbl(IIf(IsNumeric(A13), A13, 0))
    A14 = CDbl(IIf(IsNumeric(A14), A14, 0))
    A15 = CDbl(IIf(IsNumeric(A15), A15, 0))
    A16 = CDbl(IIf(IsNumeric(A16), A16, 0))
    A17 = CDbl(IIf(IsNumeric(A17), A17, 0))
    A18 = CDbl(IIf(IsNumeric(A18), A18, 0))

    For BB = 1 To ListCount
        With lstInvoice.ListItems(BB)
            A1 = CDbl(A1) + CDbl(IIf(IsNumeric(.SubItems(11)), .SubItems(11), 0))
            A2 = CDbl(A2) + CDbl(IIf(IsNumeric(.SubItems(12)), .SubItems(12), 0))
            A3 = CDbl(A3) + CDbl(IIf(IsNumeric(.SubItems(13)), .SubItems(13), 0))
            A4 = CDbl(A4) + CDbl(IIf(IsNumeric(.SubItems(14)), .SubItems(14), 0))
            A5 = CDbl(A5) + CDbl(IIf(IsNumeric(.SubItems(15)), .SubItems(15), 0))
            A6 = CDbl(A6) + CDbl(IIf(IsNumeric(.SubItems(16)), .SubItems(16), 0))
            A7 = CDbl(A7) + CDbl(IIf(IsNumeric(.SubItems(17)), .SubItems(17), 0))
            A8 = CDbl(A8) + CDbl(IIf(IsNumeric(.SubItems(18)), .SubItems(18), 0))
            A9 = CDbl(A9) + CDbl(IIf(IsNumeric(.SubItems(19)), .SubItems(19), 0))
            A10 = CDbl(A10) + CDbl(IIf(IsNumeric(.SubItems(20)), .SubItems(20), 0))
            A11 = CDbl(A11) + CDbl(IIf(IsNumeric(.SubItems(21)), .SubItems(21), 0))
            A12 = CDbl(A12) + CDbl(IIf(IsNumeric(.SubItems(22)), .SubItems(22), 0))
            A13 = CDbl(A13) + CDbl(IIf(IsNumeric(.SubItems(23)), .SubItems(23), 0))
            A14 = CDbl(A14) + CDbl(IIf(IsNumeric(.SubItems(24)), .SubItems(24), 0))
            A15 = CDbl(A15) + CDbl(IIf(IsNumeric(.SubItems(25)), .SubItems(25), 0))
            A16 = CDbl(A16) + CDbl(IIf(IsNumeric(.SubItems(26)), .SubItems(26), 0))
            A17 = CDbl(A17) + CDbl(IIf(IsNumeric(.SubItems(27)), .SubItems(27), 0))
            A18 = CDbl(A18) + CDbl(IIf(IsNumeric(.SubItems(28)), .SubItems(28), 0))
        End With
    Next


End Sub

Private Sub txtMC_KeyPress(KeyAscii As Integer)
     KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub TxtNCGPFee_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtNCMedicalRPT_KeyPress(KeyAscii As Integer)
     KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtNCMedicine_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtNCOthers_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub Cboclinic_LostFocus()
    If Mid(CboPAYto, 1, 1) = "C" Then
        CboPayee.Text = Cboclinic
    End If
End Sub

Private Sub CboPayee_LostFocus()
    CboPayee = ChkStr(CboPayee)
End Sub

Private Sub CboPAYto_LostFocus()
    If Mid(CboPAYto, 1, 1) = "C" Then
        TxtInvPayType = "Cashless"
    Else
        TxtInvPayType = "Reimbursement"
        CboPayee.Text = "Member"
    End If
End Sub

Private Sub CboPatientID_LostFocus()
    If CboPatientID.ListIndex >= 0 Then
        txtdob.Text = DOBArr(CboPatientID.ListIndex)
        TxtPatientDOB.Text = DOBArr(CboPatientID.ListIndex)
'        If IsDate(txtDateVisit) = True Then
'            Call InvoiceINFO
'        End If
    End If
End Sub

Private Sub ChkKey_Click()
    txtKeyNo = ""
    If ChkKey.Value = 1 Then
        lblkey.Caption = "Member No."
    Else
        lblkey.Caption = "Case No."
    End If
End Sub

Private Sub CmdAdd_Click()
    AddFlag = True
    EditFlag = False
        checking = False
        Call ClearInvoice
        CmdUpdate.Visible = True
        CmdUpdate.Caption = "&Save"
        CmdUpdate.Enabled = True
        CmdEdit.Visible = False
        CmdAdd.Enabled = False
        CmdTotal.Enabled = False
        checking = True
        Call EntriesEnable(True)
'        TxtAccess.Enabled = False
    
End Sub

Private Sub CmdClear_Click()
    Call ClearEntries
    lstInvoice.ListItems.Clear
    TxtInvPayType = "Cashless"
    ChkKey.Value = 1
    CboPatientID.Clear
    CboClmCat = "GP - GP Clinic Consultation"
    CboReason1.Text = ""
    CboReason2.Text = ""
    CboReason3.Text = ""
    Call EntriesEnable(False)
End Sub

Private Sub cmdEdit_Click()
    Call EntriesEnable(True)
    CmdUpdate.Visible = True
    CmdUpdate.Enabled = True
    CmdUpdate.Caption = "&Update"
    CmdEdit.Visible = False
    CmdEdit.Enabled = False
    FrameCase.Enabled = True
    EditFlag = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    frmCasList.Source = 6
    frmCasList.Show vbModal
End Sub

Private Sub CaseDet()
    Dim CISCAS As New ADODB.Recordset
    Dim PolStatus As New ADODB.Recordset
    Dim CISCasSql As String
    Dim temp As String
    temp = txtKeyNo
    Call ClearEntries
    txtKeyNo = temp
    
'    ScanCode = TxtKeyNo
'    If Mid(ScanCode, 1, 2) = "LO" Then
'        TxtKeyNo = ""
'        Call SearchForm(ScanCode)
'    Else
        If Trim(txtKeyNo) <> "" Then
            If ChkKey.Value = 1 Then 'MEMBER
                Call MBMInfo
                
            Else 'CASE NUMBER
                CISCAS.Open "select * from CISCas WHERE CASNumber = '" & Trim(txtKeyNo) & "'", MediConnCISDB, adOpenKeyset, adLockOptimistic
            
                If CISCAS.RecordCount > 0 Then
                    MembershipNo = CISCAS!MBMNumber
                                 
                    If ChkStr(CISCAS!CASStatus) = "O" Then
                        cboStatus.Text = "Open"
                    ElseIf ChkStr(CISCAS!CASStatus) = "C" Then
                        cboStatus.Text = "Close"
                    ElseIf ChkStr(CISCAS!CASStatus) = "D" Then
                        cboStatus.Text = "Delete"
                    End If
                    If CISCAS!Claimability = "A" Then
                        cboClaimability = "Accepted"
                    ElseIf CISCAS!Claimability = "R" Then
                        cboClaimability = "Repudiated"
                    ElseIf CISCAS!Claimability = "C" Then
                        cboClaimability = "Cancelled"
                    End If
                    If CISCAS!Pending = "Y" Then
                        cboPending = "Yes"
                    Else
                        cboPending = "No"
                    End If
                    Call MBMInfo
                    txtDateVisit = CISCAS!DateVisit
                    CboPatientID.Text = CISCAS!CoverID
                    Call InvoiceINFO
                    If ChkKey.Value = 0 Then
                        CmdAdd.Enabled = True
'                        CmdEdit.Enabled = True
                    End If
                Else
                    MsgBox "Record not found! ", vbCritical
                End If
            End If
        Else
            MsgBox "Please keyin Member No. or Cse Number. !", vbOKOnly + vbCritical
        End If
    'End If
End Sub


Private Sub CmdUpdate_Click()
    Call SaveRecord
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
    bExit = False
    EditFlag = False
    GenCaseNo = False
    checking = True   ' do the checking on invoice date
    Call EntriesEnable(False)
    Call ComboListing
    CmdAdd.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = True
    CmdTotal.Enabled = True
    CmdUpdate.Visible = False
    'Call DoInitialSettings
    intIndex = 0
End Sub

Private Sub ComboListing()
Dim ClinicRec As New ADODB.Recordset
Dim RepRec As New ADODB.Recordset
Dim TreatRec As New ADODB.Recordset
        
    ChkKey.Value = 1
    TxtInvPayType = "Cashless"
    ClinicRec.Open "SELECT CLNCode, CLNName FROM CLN", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until ClinicRec.EOF
        Cboclinic.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
        CboPayee.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
        ClinicRec.MoveNext
    Loop
    ClinicRec.Close
    Set ClinicRec = Nothing
    
    RepRec.Open "SELECT REPCode, RepDescription FROM REPCode", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until RepRec.EOF
        CboRepudiate1.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        CboRepudiate2.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        CboRepudiate3.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        RepRec.MoveNext
    Loop
    RepRec.Close
    Set RepRec = Nothing
    
    TreatRec.Open "SELECT ReasonCode, ReasonDesc FROM TreatmentReason", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until TreatRec.EOF
        CboReason1.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        CboReason2.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        CboReason3.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        TreatRec.MoveNext
    Loop
    TreatRec.Close
    Set RepRec = Nothing
    
    CboPAYto.AddItem "C - Clinic"
    CboPAYto.AddItem "M - Member"
    CboPAYto.ListIndex = 0
    
    CboClmCat.AddItem "GP - GP Clinic Consultation"
    CboClmCat.AddItem "SP - SP Specialist Consultation"
    CboClmCat.AddItem "OT - Others"
    
    cboClaimability.AddItem "A - Accepted"
    cboClaimability.AddItem "R - Rejected"
    cboClaimability.AddItem "C - Cancelled"
    
    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"
        
    cboStatus.AddItem "O - Open"
    cboStatus.AddItem "C - Close"
    cboStatus.AddItem "D - Delete"
End Sub
Private Sub ClearEntries()
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
        
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        End If
    Next ctl
    
    txtINvDate = ""
    Cboclinic.Text = ""
    CboPayee.Text = ""
    
    CmdAdd.Enabled = False
    CmdUpdate.Visible = False
    CmdDelete.Enabled = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
    lblBordx = ""
    lblVersion = ""
    lblLastVersion = ""
    lblWs = ""
    txtINVAmount = 0
End Sub


Private Sub lstInvoice_Click()
    If lstInvoice.ListItems.count <> 0 Then
    CmdEdit.Enabled = False
        If lstInvoice.SelectedItem.SubItems(3) = True Then
            CboPAYto.Enabled = False
        Else
            CboPAYto.Enabled = True
        End If
        Call EnableCmdButton(Me)
        CmdEdit.Enabled = True
        CmdEdit.Visible = True
        TxtIndex = lstInvoice.SelectedItem.Text
        txtInvoiceNo = lstInvoice.SelectedItem.SubItems(4)
        txtInvoiceDate = lstInvoice.SelectedItem.SubItems(5)
        Cboclinic = lstInvoice.SelectedItem.SubItems(6)
        If Trim(lstInvoice.SelectedItem.SubItems(7)) = "GP" Then
            CboClmCat = "GP - GP Clinic Consultation"
        ElseIf Trim(lstInvoice.SelectedItem.SubItems(7)) = "SP" Then
            CboClmCat = "SP - SP Specialist Consultation"
        Else
            CboClmCat = "OT - Others"
        End If
        
        CboPAYto = lstInvoice.SelectedItem.SubItems(9)
        If Mid(CboPAYto, 1, 2) = "C" Then
            TxtInvPayType = "Cashless"
        Else
            TxtInvPayType = "Reimbursement"
        End If
        CboPayee = lstInvoice.SelectedItem.SubItems(10)
        A1 = lstInvoice.SelectedItem.SubItems(13)
        A2 = lstInvoice.SelectedItem.SubItems(14)
        A3 = lstInvoice.SelectedItem.SubItems(15)
        A4 = lstInvoice.SelectedItem.SubItems(16)
        A5 = lstInvoice.SelectedItem.SubItems(17)
        A6 = lstInvoice.SelectedItem.SubItems(18)
        A7 = lstInvoice.SelectedItem.SubItems(19)
        A8 = lstInvoice.SelectedItem.SubItems(20)
        A9 = lstInvoice.SelectedItem.SubItems(21)
        A10 = lstInvoice.SelectedItem.SubItems(22)
        A11 = lstInvoice.SelectedItem.SubItems(23)
        A12 = lstInvoice.SelectedItem.SubItems(24)
        A13 = lstInvoice.SelectedItem.SubItems(25)
        A14 = lstInvoice.SelectedItem.SubItems(26)
        A15 = lstInvoice.SelectedItem.SubItems(27)
        A16 = lstInvoice.SelectedItem.SubItems(28)
        A17 = lstInvoice.SelectedItem.SubItems(29)
        A18 = lstInvoice.SelectedItem.SubItems(30)
        TxtNCGPFee = lstInvoice.SelectedItem.SubItems(31)
        TxtNCMedicine = lstInvoice.SelectedItem.SubItems(32)
        TxtNCMedicalRPT = lstInvoice.SelectedItem.SubItems(33)
        TxtNCOthers = lstInvoice.SelectedItem.SubItems(34)
        If lstInvoice.SelectedItem.SubItems(35) = "A" Then
            cboClaimability.Text = "A - Accepted"
        ElseIf lstInvoice.SelectedItem.SubItems(35) = "R" Then
            cboClaimability.Text = "R - Rejected"
        ElseIf lstInvoice.SelectedItem.SubItems(35) = "C" Then
            cboClaimability.Text = "C - Cancelled"
        End If
        If lstInvoice.SelectedItem.SubItems(36) = "Y" Then
            cboPending.Text = "Y - Yes"
        ElseIf lstInvoice.SelectedItem.SubItems(36) = "N" Then
            cboPending.Text = "N - No"
        End If
        CboReason1.Text = lstInvoice.SelectedItem.SubItems(37)
        CboReason2.Text = lstInvoice.SelectedItem.SubItems(38)
        CboReason3.Text = lstInvoice.SelectedItem.SubItems(39)
        txtMC = lstInvoice.SelectedItem.SubItems(40)
        CboRepudiate1.Text = lstInvoice.SelectedItem.SubItems(41)
        CboRepudiate2.Text = lstInvoice.SelectedItem.SubItems(42)
        CboRepudiate3.Text = lstInvoice.SelectedItem.SubItems(43)
        lblBordx = lstInvoice.SelectedItem.SubItems(44)
        lblWs = lstInvoice.SelectedItem.SubItems(45)
        lblVersion = lstInvoice.SelectedItem.SubItems(2)
        
        Call EntriesEnable(False)
        If ChkKey.Value = 1 Then
            CmdAdd.Enabled = False
            CmdEdit.Enabled = False
            
        End If
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub txtInvoiceNo_LostFocus()
    txtInvoiceNo = ChkStr(txtInvoiceNo)
End Sub

Private Sub TxtKeyNo_LostFocus()
    If Trim(txtKeyNo) <> "" Then
        Call ClearInvoice

        txtKeyNo = ChkStr(txtKeyNo)
        If ChkKey.Value = 1 Then
            MembershipNo = txtKeyNo
        Else
            CASENumber = txtKeyNo
        End If
        
        Call CaseDet
    End If
End Sub

Private Sub txtDateVisit_LostFocus()
    If txtDateVisit <> "__/__/____" Then
        If IsDate(txtDateVisit) = False Then
            MsgBox "Invalid date format!", vbCritical
        ElseIf Format(txtDateVisit, "yyyymmdd") > Format(Date, "yyyymmdd") Then
            MsgBox "Date Visit cannot greater than system date ! ", vbCritical
            txtDateVisit.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceDate_LostFocus()
    If txtInvoiceDate <> "__/__/____" Then
        If IsDate(txtInvoiceDate) = False Then
            MsgBox "Invalid date format!", vbCritical
        ElseIf Format(txtInvoiceDate, "yyyymmdd") > Format(Date, "yyyymmdd") Then
            MsgBox "Invoice date cannot greater than system date ! ", vbCritical
            txtInvoiceDate.SetFocus
        End If
    End If
End Sub

Private Sub MBMInfo()
Dim MBMRec As New ADODB.Recordset
Dim MBMSql As String
Dim COVRec As New ADODB.Recordset
Dim COVSql As String
Dim Cnt As Integer

    If Len(Trim(MembershipNo)) Then
        MBMSql = "SELECT MBMNumber, MBMName, MBMPolicyNo, MBMPayorEffDate, MBMPayorExpDate, PAYCode, " & _
                " MBMPolicyNo,MBMStatus, MBMCOVID,INSCode, PLNCode, MBMDOB FROM MBM WHERE MBMNumber='" & Trim(MembershipNo) & " '"
        MBMRec.Open MBMSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        CboPatientID.Clear
        If MBMRec.RecordCount <= 0 Then
            MsgBox "Membership No not found ! ", vbOKOnly + vbCritical, DialogBoxTitle
            txtKeyNo.SetFocus
        Else
            CboPatientID.AddItem MBMRec!MBMCOVID & " - " & MBMRec!MBMName
            DOBArr(0) = MBMRec!MBMDOB
            txtTabMBMNo = MBMRec!MBMNumber
            TxtPayor = MBMRec!PAYCode
            txtPlanCode = MBMRec!PLNCode
            TxtInsType = MBMRec!INSCode
            txtMBMName = MBMRec!MBMName
            txtTabMBMName = MBMRec!MBMName
            txtEffDate = MBMRec!MBMPayorEffDate
            txtExpDate = MBMRec!MBMPayorExpDate
            TxtPolNo = MBMRec!MBMPolicyNo
            txtPolStatus = MBMRec!MBMStatus
            CboPatientID.Enabled = True
            If Trim(MBMRec!INSCode) = "I" Or Trim(MBMRec!INSCode) = "G" Then
                TxtPatientDOB.Text = MBMRec!MBMDOB
                txtdob.Text = MBMRec!MBMDOB
                CboPatientID.Text = MBMRec!MBMCOVID & " - " & MBMRec!MBMName
                CboPatientID.Enabled = False
            End If
            
            COVSql = "SELECT MBMCName, MBMCCoverID, MBMCDOB FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & " '"
            COVRec.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If COVRec.RecordCount Then
                Do Until COVRec.EOF
                    Cnt = Cnt + 1
                    CboPatientID.AddItem COVRec!MBMCCoverID & " - " & COVRec!MBMCName
                    DOBArr(Cnt) = COVRec!MBMCDOB
                    COVRec.MoveNext
                Loop
            End If
            COVRec.Close
            Set COVRec = Nothing
            If Date - CDate(txtExpDate) > 14 Then
            LblExpDate.Caption = "Expired"
        End If
        End If
        MBMRec.Close
        Set MBMRec = Nothing
        
    End If
End Sub

Private Function GenerateCISCaseNo(PayorCode)
    Dim CISCAQ As New ADODB.Recordset
    Dim strsql As String
    Dim CISSeqNo As String
    Dim CAQCharacter As String
    
    strsql = "Select PAYCode, CISSeqNo from CISCAQ Where PAYCode = '" & PayorCode & "'"
    CISCAQ.Open strsql, MediConnCISDB, adOpenStatic, adLockOptimistic
    
    If CISCAQ.RecordCount <= 0 Then
        With CISCAQ
            .AddNew
            !PAYCode = PayorCode
            !CISSeqNo = "000001"
            .Update
            CISSeqNo = "000001"
        End With
    Else
        CISCAQ!CISSeqNo = Format(CISCAQ!CISSeqNo + 1, "000000")
        CISSeqNo = CISCAQ!CISSeqNo
        CISCAQ.Update
    End If

    GenerateCISCaseNo = PayorCode & "C" & CISSeqNo
    CISCAQ.Close
    Set CISCAQ = Nothing
End Function
Private Sub getVersion()
    Dim Version As New ADODB.Recordset
    Dim strsql As String
    strsql = "select casNumber, max(clmversion) AS lastVersionNo from claim where casNUmber ='" & Trim(CASENumber) & "' " & _
            " and ClaimStatus ='C' " & _
            "group by casNumber"
    Version.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    If Not Version.EOF Then
        lblLastVersion = Version!lastVersionNo + 1
    Else
        lblLastVersion = 1
    End If
    Version.Close
    Set Version = Nothing
End Sub

Private Sub InvoiceINFO()
Dim CISCAS1 As New ADODB.Recordset
Dim CASSql, CLMInvSql As String
Dim rst2 As New ADODB.Recordset
Dim tempMBM As String
Dim RecList As ListItem
Dim ClmInvACt As New ADODB.Recordset
Dim CLmInvACTSql As String
Dim strsql As String
Dim SSS As Integer

    lstInvoice.ListItems.Clear
    If ChkKey.Value = 0 Then
        strsql = " WHERE CASNumber='" & CASENumber & "'"
    Else
        strsql = " WHERE MBMNumber='" & MembershipNo & "' AND COVERID='" & Mid(CboPatientID, 1, 2) & "' AND DateVisit='" & Format(txtDateVisit, "MM/DD/YYYY") & "'"
    End If
    CASSql = "SELECT MBMNumber, * FROM View_Clm_Invoice " & strsql
    CISCAS1.Open CASSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    'SSS = CISCAS1.RecordCount
    Do Until CISCAS1.EOF
        With CISCAS1
            Set RecList = lstInvoice.ListItems.Add(, , !InvINdex)
                RecList.SubItems(1) = !CASNumber
                RecList.SubItems(2) = IIf(Len(!CLMVersion), !CLMVersion, "")
                RecList.SubItems(3) = !INVWSStatus
                RecList.SubItems(4) = !INVNo
                RecList.SubItems(5) = !InvDate
                RecList.SubItems(6) = !ClinicCode
                RecList.SubItems(7) = !ClaimType
                RecList.SubItems(8) = IIf(Len(!INVAmount), !INVAmount, 0)
                RecList.SubItems(9) = !INVPayTo
                RecList.SubItems(10) = !INVPayee
                RecList.SubItems(11) = IIf(Len(!ClaimType), !ClaimType, "")
                RecList.SubItems(12) = IIf(Len(!INVEnteredBy), !INVEnteredBy, "")
                RecList.SubItems(13) = IIf(Len(!CLMGPConsAmt), !CLMGPConsAmt, 0)
                RecList.SubItems(14) = IIf(Len(!CLMSPConsAmt), !CLMSPConsAmt, 0)
                RecList.SubItems(15) = IIf(Len(!CLMMedicine), !CLMMedicine, 0)
                RecList.SubItems(16) = IIf(Len(!CLMProcedures), !CLMProcedures, 0)
                RecList.SubItems(17) = IIf(Len(!CLMLab), !CLMLab, 0)
                RecList.SubItems(18) = IIf(Len(!CLMXray), !CLMXray, 0)
                RecList.SubItems(19) = IIf(Len(!CLMMRI), !CLMMRI, 0)
                RecList.SubItems(20) = IIf(Len(!CLMDental), !CLMDental, 0)
                RecList.SubItems(21) = IIf(Len(!CLMChineseMed), !CLMChineseMed, 0)
                RecList.SubItems(22) = IIf(Len(!CLMTraditionalMEd), !CLMTraditionalMEd, 0)
                RecList.SubItems(23) = IIf(Len(!CLMphysioTheraphy), !CLMphysioTheraphy, 0)
                RecList.SubItems(24) = IIf(Len(!CLMAccident), !CLMAccident, 0)
                RecList.SubItems(25) = IIf(Len(!CLMPRVCheckup), !CLMPRVCheckup, 0)
                RecList.SubItems(26) = IIf(Len(!CLMVision), !CLMVision, 0)
                RecList.SubItems(27) = IIf(Len(!CLMID), !CLMID, 0)
                RecList.SubItems(28) = IIf(Len(!CLMTSSTO), !CLMTSSTO, 0)
                RecList.SubItems(29) = IIf(Len(!CLMOTHERS1), !CLMOTHERS1, 0)
                RecList.SubItems(30) = IIf(Len(!CLMOTHERS2), !CLMOTHERS2, 0)
                RecList.SubItems(31) = IIf(Len(!GPFee), !GPFee, 0)
                RecList.SubItems(32) = IIf(Len(!Medicine), !Medicine, 0)
                RecList.SubItems(33) = IIf(Len(!MedicalRpt), !MedicalRpt, 0)
                RecList.SubItems(34) = IIf(Len(!Others), !Others, 0)
                RecList.SubItems(35) = !Claimability
                RecList.SubItems(36) = !Pending
                RecList.SubItems(37) = IIf(Len(!CASTreatReason1), !CASTreatReason1, "")
                RecList.SubItems(38) = IIf(Len(!CASTreatReason2), !CASTreatReason2, "")
                RecList.SubItems(39) = IIf(Len(!CASTreatReason3), !CASTreatReason3, "")
                RecList.SubItems(40) = IIf(Len(!MCDays), !MCDays, 0)
                RecList.SubItems(41) = IIf(Len(!CASREPCode1), !CASREPCode1, "")
                RecList.SubItems(42) = IIf(Len(!CASREPCode2), !CASREPCode2, "")
                RecList.SubItems(43) = IIf(Len(!CASREPCode3), !CASREPCode3, "")
                RecList.SubItems(44) = IIf(Len(!INVBordxed), !INVBordxed, "")
                RecList.SubItems(45) = IIf(Len(!INVWSStatus), !INVWSStatus, "")
                CISCAS1.MoveNext
            End With
    Loop
    
    CISCAS1.Close
    Set CISCAS1 = Nothing
    Call getVersion
End Sub

Private Sub SaveRecord()
Dim RecordSaved
Dim CISCasRec As New ADODB.Recordset
Dim CISCasSql As String
Dim startTrans As Boolean

    If Trim(txtKeyNo) = "" Then
        MsgBox "Please keyin Member No or Case No. ! ", vbCritical
        txtKeyNo.SetFocus
'    ElseIf Trim(txtMBMName) = "" Then
'        MsgBox "Please select a record before proceeding! ", vbCritical
'        txtCaseNo.SetFocus
    ElseIf txtDateVisit = "__/__/____" Or txtDateVisit = "" Then
        MsgBox "Please Enter Date of Visit! ", vbCritical
    ElseIf Trim(CboPAYto) = "" Then
        MsgBox "Please select who to be paid (Member / Clinic )! ", vbCritical
        CboPAYto.SetFocus
    ElseIf Trim(CboPayee) = "" Then
        MsgBox "Please keyin the person/Clinic to be paid! ", vbCritical
        CboPayee.SetFocus
    ElseIf Trim(Cboclinic) = "" Then
        MsgBox "Please keyin the Clinic code ! ", vbCritical
        Cboclinic.SetFocus
    ElseIf Trim(txtInvoiceNo) = "" Then
        MsgBox "Please keyin Invoice No.! ", vbCritical
        txtInvoiceNo.SetFocus
    ElseIf txtInvoiceDate = "__/__/____" Or txtInvoiceDate = "" Then
        MsgBox "Please key in Invoice date! ", vbCritical
    ElseIf Trim(CboClmCat) = "" Then
        MsgBox "Please SELECT claim Category ! ", vbCritical
        cboHosCode.SetFocus
    Else
    
        'If GenCaseNo = False Then
        If ChkKey.Value = 1 Then
            CISCasSql = "SELECT * FROM View_Clm_Invoice where MBMNumber = '" & Trim(txtKeyNo) & _
                      "' AND COVERID ='" & Mid(CboPatientID, 1, 2) & "' AND DateVisit = '" & txtDateVisit & "'"
        Else
            CISCasSql = "SELECT * FROM View_Clm_Invoice where CASNumber = '" & Trim(txtKeyNo) & _
                      "' AND COVERID ='" & Mid(CboPatientID, 1, 2) & "' AND DateVisit = '" & txtDateVisit & _
                      "' AND InvNo = '" & txtInvoiceNo & "'"
        End If
                
        CISCasRec.Open CISCasSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        If CISCasRec.RecordCount And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            CISCasRec.Close
            If ChkKey.Value = 1 Then
                txtDateVisit.SetFocus
            Else
                txtInvoiceNo.SetFocus
            End If
            Set CISCasRec = Nothing
           
            Exit Sub
        Else
                
            RecordSaved = MsgBox("Are you sure", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                MediConnCISWS.beginTrans
                startTrans = True
               
                Call getTotal
                Call SaveCase
                Call SaveInvoice
                Call saveActual
                Call SaveReason
               ' Call saveActualOthers
                Call saveNotCover
            
                MediConnCISWS.CommitTrans
                startTrans = False
                If ChkKey.Value = 1 Then
                   MsgBox "Case Number : " & NewCaseNo, vbCritical
                Else
                    If CmdUpdate.Caption = "&Update" Then
                        MsgBox "Record Updated ! ", vbOKOnly
                    Else
                        MsgBox "Invoice Added ! ", vbOKOnly
                    End If
                End If
                Call EnableCmdButton(Me)
                
                CmdClear.Enabled = True
                Call ClearInvoice
'                ChkKey.value = 1
                Call EntriesEnable(False)
                Call InvoiceINFO
                CmdTotal.Enabled = True
                CmdEdit.Visible = True
                CmdEdit.Enabled = False
                EditFlag = False
            End If
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        MediConnCISWS.RollbackTrans
    End If
    MsgBox Err.Description
End Sub

Private Function getTotal() As Double
    Dim ctl  As Control
    txtINVAmount = 0
    For Each ctl In Me.Controls
        If Mid(ctl.Name, 1, 1) = "A" And Len(ctl.Name) <= 7 Then
            If IsNumeric(ctl.Text) = True Then
                txtINVAmount = txtINVAmount + CDbl(ctl.Text)
            End If
        End If
    Next
End Function
Private Sub SaveInvoice()
Dim ClaimInvoice  As New ADODB.Recordset
    
    If EditFlag = False Then
        ClaimInvoice.Open "Select INVINdex , * from ClaimInvoice", MediConnCISWS, adOpenKeyset, adLockOptimistic
        ClaimInvoice.AddNew
        With ClaimInvoice
            !CLMVersion = lblLastVersion
            !INVBordxed = 0
            !INVWSStatus = 0
            !INVEnteredDate = Date
            !INVEnteredBy = Logon
        End With
    Else
        strsql = "Select INVINdex , * from ClaimInvoice" & _
            " where INVIndex ='" & Trim(TxtIndex) & "'"
        ClaimInvoice.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    End If
    
    With ClaimInvoice
        !INVNo = txtInvoiceNo
        !CASNumber = NewCaseNo
        !InvDate = txtInvoiceDate
        !ClaimType = Mid(CboClmCat, 1, 2)
        !MCDays = IIf(IsNumeric(txtMC), txtMC, 0)
        !INVLastUpdateBy = Logon
        !INVLastUpdateDate = Date
        If IsNumeric(txtINVAmount) = True Then
            !INVAmount = txtINVAmount
        Else
            !INVAmount = 0
        End If
        !INVPayee = Trim(Left(CboPayee, IIf(InStr(1, CboPayee, "-") - 1 < 0, Len(CboPayee), InStr(1, CboPayee, "-") - 1)))
        !INVPAYType = Mid(TxtInvPayType, 1, 1)
        !INVPayTo = Mid(CboPAYto, 1, 1)

    End With
    ClaimInvoice.Update
    
    If EditFlag = False Then
        Dim ClaimInvoiceLastIndex  As New ADODB.Recordset
        ClaimInvoiceLastIndex.Open "Select INVINdex from ClaimInvoice Order by INVIndex DESC", MediConnCISWS, adOpenKeyset, adLockOptimistic
        If Not ClaimInvoiceLastIndex.EOF Then
            intIndex = ClaimInvoiceLastIndex!InvINdex
        Else
            intIndex = 1
        End If
        ClaimInvoiceLastIndex.Close
        Set ClaimInvoiceLastIndex = Nothing
    Else
        intIndex = TxtIndex
    End If
    
    ClaimInvoice.Close
    Set ClaimInvoice = Nothing
End Sub
Private Sub saveActual()
    Dim ClaimActual As New ADODB.Recordset
    Dim strsql As String
 
    If EditFlag = False Then
        ClaimActual.Open "Select INVINdex , * from ClaimInvoiceActual", MediConnCISWS, adOpenKeyset, adLockOptimistic
        ClaimActual.AddNew
        ClaimActual!InvINdex = intIndex
    Else
        strsql = "Select INVINdex , * from ClaimInvoiceActual where INVIndex =" & intIndex
        ClaimActual.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    End If
    
    With ClaimActual
        !CLMGPConsAmt = IIf(IsNumeric(A1), A1, 0)
        !CLMSPConsAmt = IIf(IsNumeric(A2), A2, 0)
        !CLMMedicine = IIf(IsNumeric(A3), A3, 0)
        !CLMProcedures = IIf(IsNumeric(A4), A4, 0)
        !CLMLab = IIf(IsNumeric(A5), A5, 0)
        !CLMXray = IIf(IsNumeric(A6), A6, 0)
        !CLMMRI = IIf(IsNumeric(A7), A7, 0)
        !CLMDental = IIf(IsNumeric(A8), A8, 0)
        !CLMChineseMed = IIf(IsNumeric(A9), A9, 0)
        !CLMTraditionalMEd = IIf(IsNumeric(A10), A10, 0)
        !CLMphysioTheraphy = IIf(IsNumeric(A11), A11, 0)
        !CLMAccident = IIf(IsNumeric(A12), A12, 0)
        !CLMPRVCheckup = IIf(IsNumeric(A13), A13, 0)
        !CLMVision = IIf(IsNumeric(A14), A14, 0)
        !CLMID = IIf(IsNumeric(A15), A15, 0)
        !CLMTSSTO = IIf(IsNumeric(A16), A16, 0)
        !CLMOTHERS1 = IIf(IsNumeric(A17), A17, 0)
        !CLMOTHERS2 = IIf(IsNumeric(A18), A18, 0)
        !CLMOthers1Desc = txtOtherDesc1
        !CLMOthers2Desc = txtOtherDesc2
    End With
    ClaimActual.Update
    
    ClaimActual.Close
    Set ClaimActual = Nothing
End Sub

Private Sub saveNotCover()
Dim ClaimNotCover  As New ADODB.Recordset
Dim strsql As String
    If EditFlag = False Then
        ClaimNotCover.Open "Select INVINdex , * from ClaimInvoiceNotCover ", MediConnCISWS, adOpenKeyset, adLockOptimistic
        ClaimNotCover.AddNew
        ClaimNotCover!InvINdex = intIndex
    Else
        strsql = "Select INVINdex , * from ClaimInvoiceNotCover where INVIndex =" & intIndex
        ClaimNotCover.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    End If
    With ClaimNotCover
        !GPFee = IIf(Len(TxtNCGPFee), TxtNCGPFee, 0)
        !Medicine = IIf(Len(TxtNCMedicine), TxtNCMedicine, 0)
        !MedicalRpt = IIf(Len(TxtNCMedicalRPT), TxtNCMedicalRPT, 0)
        !Others = IIf(Len(TxtNCOthers), TxtNCOthers, 0)
    End With
    ClaimNotCover.Update
    
    ClaimNotCover.Close
    Set ClaimNotCover = Nothing
End Sub




Private Sub SaveCase()
Dim CISCAS  As New ADODB.Recordset
Dim CISSql As String
    
    Hospitals = ""
    
    If ChkKey.Value = 1 Then
        CISSql = "SELECT * FROM CISCAS "
        CISCAS.Open CISSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISCAS.AddNew
        NewCaseNo = GenerateCISCaseNo(TxtPayor)
    ElseIf ChkKey.Value = 0 Then
        CISSql = "SELECT * FROM CISCAS where CASNumber = '" & Trim(CASENumber) & "'"
        CISCAS.Open CISSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        NewCaseNo = CASENumber
    End If
    
    
    With CISCAS
        If ChkKey.Value = 1 Then
            !CASNumber = NewCaseNo
            !MBMNumber = txtKeyNo
            !CoverID = Mid(CboPatientID, 1, 2)
            !DateVisit = txtDateVisit
            !CASStatus = "O"
            !Claimability = "A"
            !Pending = "N"
            !ClinicCode = Trim(Left(Cboclinic, IIf(InStr(1, Cboclinic, "-") - 1 < 0, Len(Cboclinic), InStr(1, Cboclinic, "-") - 1)))
            !REgisterDate = Date
            !REgisterBy = Logon
        Else
            !CASStatus = Mid(cboStatus, 1, 1)
            !Claimability = Mid(cboClaimability, 1, 1)
            !Pending = Mid(cboPending, 1, 1)
            !CASTreatReason1 = Trim(Left(CboReason1, IIf(InStr(1, CboReason1, "-") - 1 < 0, Len(CboReason1), InStr(1, CboReason1, "-") - 1)))
            !CASTreatReason2 = Trim(Left(CboReason2, IIf(InStr(1, CboReason2, "-") - 1 < 0, Len(CboReason2), InStr(1, CboReason2, "-") - 1)))
            !CASTreatReason3 = Trim(Left(CboReason3, IIf(InStr(1, CboReason3, "-") - 1 < 0, Len(CboReason3), InStr(1, CboReason3, "-") - 1)))
        End If
        If Mid(cboClaimability, 1, 1) = "R" Then
            Call SaveRepudiation
        End If
            !Remarks = TxtRemarks
    End With
    CISCAS.Update
    
    CISCAS.Close
    Set CISCAS = Nothing
End Sub
Private Sub SaveRepudiation()
Dim CASREP As New ADODB.Recordset
Dim CAS As New ADODB.Recordset
Dim strsql As String
    
    CASREP.Open "select * from CISCASRepudiation", MediConnCISDB, adOpenKeyset, adLockOptimistic
        If CASREP.RecordCount = 0 Then
            CASREP.AddNew
            
        End If
        
        With CASREP
            
            !CASNumber = CASENumber
            !CASREPCode1 = Trim(Mid(ChkStr(CboRepudiate1), 1, 7))
            !CASREPCode2 = Trim(Mid(ChkStr(CboRepudiate2), 1, 7))
            !CASREPCode3 = Trim(Mid(ChkStr(CboRepudiate2), 1, 7))
            
            !CASREPIssuedBy = Logon
            !CASREPDate = Date
        End With
        CASREP.Update
       
    CASREP.Close
    Set CASREP = Nothing
End Sub

Private Sub EntriesEnable(status As Boolean)
    FrameCLMCat.Enabled = status
    FrameInvoice.Enabled = status
    FrameCharges.Enabled = status
    FrameReason.Enabled = status
    FrameNCover.Enabled = status
    FrameMBMInfo.Enabled = status
    FrameCASeInfo.Enabled = status
    txtMC.Enabled = status
End Sub

Private Sub ClearInvoice()
On Error Resume Next
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.TabIndex >= 7 And ctl.TabIndex <= 41 Then
            If TypeOf ctl Is TextBox Then
                If ctl.Name <> "txtCaseNo" Then
                    ctl.Text = ""
                End If
            ElseIf TypeOf ctl Is ComboBox Then
                 'If ctl.Name = "CboPayType" Then
                 '   CboPayType.ListIndex = 1
                'Else
                '    ctl.Text = ""
                'End If
            ElseIf TypeOf ctl Is MaskEdBox Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
        End If
    Next ctl
    TxtInvPayType = "Cashless"
    CboPAYto.Text = "C - Clinic"
    Cboclinic.Text = ""
    CboPayee.Text = ""
    CboReason1.Text = ""
    CboReason2.Text = ""
    CboReason3.Text = ""
    CmdUpdate.Visible = False
    CmdDelete.Enabled = False
    CmdTotal.Enabled = True
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
End Sub

Private Sub SaveReason()
'Dim ClmTreatReason As New ADODB.Recordset
'Dim strsql As String

'    If EditFlag = False Then
'        ClmTreatReason.Open "Select INVINdex , * from ClaimInvoiceTreatmentReason ", MediConnCISWS, adOpenKeyset, adLockOptimistic
'        ClmTreatReason.AddNew
'        ClmTreatReason!InvINdex = intIndex
'    Else
'        strsql = "Select INVINdex , * from ClaimInvoiceTreatmentReason where INVIndex =" & intIndex
'        ClmTreatReason.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
'    End If
'    With ClmTreatReason
'        !CASNumber = NewCaseNo
'        !CLMReason1 = Trim(Left(CboReason1, IIf(InStr(1, CboReason1, "-") - 1 < 0, Len(CboReason1), InStr(1, CboReason1, "-") - 1)))
'        !CLMReason2 = Trim(Left(CboReason2, IIf(InStr(1, CboReason2, "-") - 1 < 0, Len(CboReason2), InStr(1, CboReason2, "-") - 1)))
'        !CLMReason3 = Trim(Left(CLMReason3, IIf(InStr(1, CLMReason3, "-") - 1 < 0, Len(CLMReason3), InStr(1, CLMReason3, "-") - 1)))
'    End With
'    ClmTreatReason.Update
'
'    ClmTreatReason.Close
'    Set ClmTreatReason = Nothing
End Sub

