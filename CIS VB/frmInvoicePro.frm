VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInvoice 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Invoice Processing"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   120
      TabIndex        =   96
      Top             =   480
      Width           =   10815
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1080
         TabIndex        =   97
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label14 
         Caption         =   "Case No."
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   1815
      Left            =   120
      TabIndex        =   61
      Top             =   5160
      Width           =   11295
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1920
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1920
         TabIndex        =   63
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1920
         TabIndex        =   64
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1920
         TabIndex        =   65
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   1920
         TabIndex        =   66
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1920
         TabIndex        =   67
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   6000
         TabIndex        =   68
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   6000
         TabIndex        =   69
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   6000
         TabIndex        =   70
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   6000
         TabIndex        =   71
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   6000
         TabIndex        =   72
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   6000
         TabIndex        =   73
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   10080
         TabIndex        =   74
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   10080
         TabIndex        =   75
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   10080
         TabIndex        =   76
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   10080
         TabIndex        =   77
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   10080
         TabIndex        =   78
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "General Prac Cons Fee"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Special Prac Cons Fee"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Medication"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Procedures"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "Lab"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label23 
         Caption         =   "X-Ray"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "MRI"
         Height          =   255
         Left            =   4200
         TabIndex        =   89
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "Dental"
         Height          =   255
         Left            =   4200
         TabIndex        =   88
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Chinese Medicine"
         Height          =   255
         Left            =   4200
         TabIndex        =   87
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Traditional Medicine"
         Height          =   255
         Left            =   4200
         TabIndex        =   86
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label28 
         Caption         =   "Physio Theraphy"
         Height          =   255
         Left            =   4200
         TabIndex        =   85
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label29 
         Caption         =   "Accident"
         Height          =   255
         Left            =   4320
         TabIndex        =   84
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label30 
         Caption         =   "Preventive Checkup"
         Height          =   255
         Left            =   8400
         TabIndex        =   83
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "Vision"
         Height          =   255
         Left            =   8400
         TabIndex        =   82
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "I&&D"
         Height          =   255
         Left            =   8400
         TabIndex        =   81
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label33 
         Caption         =   "T&&S / STO"
         Height          =   255
         Left            =   8400
         TabIndex        =   80
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label34 
         Caption         =   "Others"
         Height          =   255
         Left            =   8400
         TabIndex        =   79
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   7440
      Width           =   11655
      Begin VB.CommandButton CmdClaerInv 
         Caption         =   "Clear &Inv"
         Height          =   330
         Left            =   8880
         TabIndex        =   39
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9720
         TabIndex        =   38
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10560
         TabIndex        =   37
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdTotal 
         Caption         =   "&Total"
         Enabled         =   0   'False
         Height          =   330
         Left            =   0
         TabIndex        =   36
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Enabled         =   0   'False
         Height          =   330
         Left            =   720
         TabIndex        =   35
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         TabIndex        =   34
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2280
         TabIndex        =   33
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   3120
         TabIndex        =   32
         Top             =   200
         Width           =   855
      End
      Begin VB.Label lblWs 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   8280
         TabIndex        =   47
         Top             =   200
         Width           =   495
      End
      Begin VB.Label Label63 
         Caption         =   "WS?"
         Height          =   255
         Left            =   7800
         TabIndex        =   46
         Top             =   200
         Width           =   495
      End
      Begin VB.Label lblBordx 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   7320
         TabIndex        =   45
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label50 
         Caption         =   "Version"
         Height          =   255
         Left            =   4680
         TabIndex        =   44
         Top             =   200
         Width           =   615
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   43
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label58 
         Caption         =   "Bordx?"
         Height          =   255
         Left            =   6600
         TabIndex        =   42
         Top             =   200
         Width           =   615
      End
      Begin VB.Label lblLastVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   41
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label54 
         Caption         =   "of"
         Height          =   255
         Left            =   5760
         TabIndex        =   40
         Top             =   200
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   11655
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   7200
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   1200
         TabIndex        =   23
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label16 
         Caption         =   "Days"
         Height          =   255
         Left            =   7920
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Medical Certificate (MC)"
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Others (specify)"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Reason"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1215
      Left            =   6960
      TabIndex        =   2
      Top             =   2280
      Width           =   4815
      Begin VB.TextBox Text9 
         Height          =   645
         Left            =   2160
         TabIndex        =   19
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Others (Specify)"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Specialist Consltation"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "GP Clinic Consultation"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2055
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
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   6735
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   4440
         TabIndex        =   100
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   99
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   820
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   315
         Left            =   4800
         TabIndex        =   60
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3840
         TabIndex        =   28
         Top             =   1080
         Width           =   2895
      End
      Begin MSMask.MaskEdBox MskBillDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   101
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Member No"
         Height          =   255
         Left            =   240
         TabIndex        =   104
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Patient"
         Height          =   255
         Left            =   3480
         TabIndex        =   103
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Date Visit"
         Height          =   255
         Left            =   240
         TabIndex        =   102
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Payee"
         Height          =   255
         Left            =   2640
         TabIndex        =   27
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Cash Type"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Pay To"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7815
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5640
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   1320
         TabIndex        =   48
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5640
         TabIndex        =   55
         Top             =   480
         Width           =   2055
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   315
         Left            =   5640
         TabIndex        =   53
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   315
         Left            =   3360
         TabIndex        =   51
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3360
         TabIndex        =   59
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1320
         TabIndex        =   57
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label40 
         Caption         =   "ClmType"
         Height          =   255
         Left            =   2640
         TabIndex        =   58
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "Nature of CLM"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label38 
         Caption         =   "Pol. Status"
         Height          =   255
         Left            =   4680
         TabIndex        =   54
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Patient DOB"
         Height          =   255
         Left            =   4680
         TabIndex        =   52
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   2655
         TabIndex        =   50
         Top             =   555
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Policy No"
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Member Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   480
      Left            =   120
      TabIndex        =   30
      Top             =   6960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   847
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
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
      NumItems        =   79
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "WS "
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "WS No"
         Object.Width           =   883
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Inv No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Inv Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "INV Amount"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Pay To"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Claim Type"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Fol. Up Date"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "DailyRB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "RB Rate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "ICUDay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ICU Rate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "HSS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "OT"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "PHService"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "IHPysVisit"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "PHTreatment"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "MR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Phone Charges"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "LodgerDay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Lodger Rate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Admin"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Tax"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Others"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "CashDay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Cash Rate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "XRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Lab"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "pre Others"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "SFee Amount"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "AFee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "EOut-Patient"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Ambfee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "OPT"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "MonthlyKidney"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "Pre60"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "PreGP"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "PreMedic"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "PreMR "
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "PreSpec"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "Post_Host"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "NotSame"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "Others"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "SurgicalPhy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "SurgicalPost"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "Others1Desc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "Others2Desc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "Others2Value"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "CLMCashAmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "CLMGuardianAmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "BoardAmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "ICUAmnt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "KidneyMonth"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "Surgicalfee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "PreNSHosp"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "In Hospital Phy Visit"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Text            =   "Organ Transplant"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "RBTax"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Text            =   "ICU Tax"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   63
         Text            =   "Lodger Tax"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "StayType"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "PayType"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "Prepared By"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "Prepared Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "MRI"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   69
         Text            =   "MXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   70
         Text            =   "MLab"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   71
         Text            =   "MPreOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   72
         Text            =   "MPreHosAmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   73
         Text            =   "Pre-Hosp Specialist Fee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   74
         Text            =   "Post Treat"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   75
         Text            =   "AccPostAmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   76
         Text            =   "NewRBTax"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   77
         Text            =   "NewICUTax"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   78
         Text            =   "NewLodgerTax"
         Object.Width           =   2540
      EndProperty
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
      Height          =   435
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
