VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceiptPayor 
   Caption         =   "Receipt From Grp Company"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Receipt from Group Company"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Top             =   45
         Width           =   3885
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   120
      TabIndex        =   55
      Top             =   5520
      Width           =   11655
      Begin VB.TextBox TxtReceiptNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "FrmReceiptPayor.frx":0000
         Left            =   6000
         List            =   "FrmReceiptPayor.frx":0028
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtReceiptAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7800
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   240
         Width           =   3705
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "FrmReceiptPayor.frx":0053
         Left            =   6000
         List            =   "FrmReceiptPayor.frx":00DB
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   500
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtChequeNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   20
         Top             =   500
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   4800
         TabIndex        =   62
         Top             =   550
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   4800
         TabIndex        =   61
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Rec No"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
         Height          =   255
         Left            =   7200
         TabIndex        =   58
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Rec Amt"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   57
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   2400
         TabIndex        =   56
         Top             =   550
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstBordx 
      Height          =   2565
      Left            =   120
      TabIndex        =   65
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bordx Amt"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Receipt Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "OS Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Receipt Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bordx Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Subsidiary"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ReceiptNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Client PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Medix Receipt No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Insured"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstWorksheet 
      Height          =   2520
      Left            =   120
      TabIndex        =   66
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4445
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Worksheet No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bordx Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Check Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "OS Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Receipt No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Receipt Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Rec. Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Bordx Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "GroupCompany"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Subsidiary"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Client PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Medix Receipt No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Insured"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   120
      TabIndex        =   38
      Top             =   360
      Width           =   11655
      Begin VB.OptionButton OptWks 
         Caption         =   "By Worksheet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   200
         Width           =   1815
      End
      Begin VB.OptionButton OptBordx 
         Caption         =   "By Bordx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   0
         Top             =   200
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtRECNoFr 
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   9
         Top             =   510
         Width           =   1215
      End
      Begin VB.TextBox txtRECNoTo 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   10
         Top             =   510
         Width           =   1575
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5160
         TabIndex        =   6
         Top             =   820
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtWSNoFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   2
         Top             =   500
         Width           =   1335
      End
      Begin VB.TextBox txtWSNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   3
         Top             =   500
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtRECDateFr 
         Height          =   315
         Left            =   8160
         TabIndex        =   11
         Top             =   750
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRECDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   12
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   4
         Top             =   750
         Width           =   1335
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   5
         Top             =   750
         Width           =   1335
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1000
         Width           =   3255
      End
      Begin VB.ComboBox CboWSStatus 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1290
         Width           =   3255
      End
      Begin VB.ComboBox CboBordxStatus 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1290
         Width           =   3255
      End
      Begin VB.ComboBox cboSubsidiary 
         Height          =   315
         ItemData        =   "FrmReceiptPayor.frx":01E7
         Left            =   1800
         List            =   "FrmReceiptPayor.frx":01E9
         Sorted          =   -1  'True
         TabIndex        =   68
         Top             =   1560
         Width           =   3255
      End
      Begin VB.ComboBox cisdbcbopayor 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   1800
         Width           =   3255
      End
      Begin VB.ComboBox cboinsuredclient 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   2090
         Width           =   3255
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmReceiptPayor.frx":01EB
         Left            =   8160
         List            =   "FrmReceiptPayor.frx":0213
         TabIndex        =   13
         Top             =   1035
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmReceiptPayor.frx":023E
         Left            =   9840
         List            =   "FrmReceiptPayor.frx":0266
         TabIndex        =   14
         Top             =   1005
         Width           =   1575
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmReceiptPayor.frx":0291
         Left            =   8160
         List            =   "FrmReceiptPayor.frx":0319
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmReceiptPayor.frx":0425
         Left            =   9840
         List            =   "FrmReceiptPayor.frx":04AD
         TabIndex        =   16
         Top             =   1275
         Width           =   1575
      End
      Begin VB.TextBox ClientPVNofrom 
         Height          =   315
         Left            =   8160
         TabIndex        =   89
         Top             =   1600
         Width           =   1215
      End
      Begin VB.TextBox ClientPVNoto 
         Height          =   315
         Left            =   9840
         TabIndex        =   90
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox MedixReceiptNofrom 
         Height          =   315
         Left            =   8160
         TabIndex        =   92
         Top             =   1890
         Width           =   1215
      End
      Begin VB.TextBox MedixReceiptNoto 
         Height          =   315
         Left            =   9840
         TabIndex        =   93
         Top             =   1850
         Width           =   1575
      End
      Begin VB.ComboBox cboInsured 
         Height          =   315
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   2180
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Insured"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   95
         Top             =   2185
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   9480
         TabIndex        =   94
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9480
         TabIndex        =   91
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Medix Receipt No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   88
         Top             =   1900
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Client PV No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6600
         TabIndex        =   87
         Top             =   1600
         Width           =   1215
      End
      Begin VB.Label lblinsurclient 
         Caption         =   "Insured Client"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   2090
         Width           =   1335
      End
      Begin VB.Label lblpayor 
         Caption         =   "Payor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Subsidiary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Recp Date from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   54
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr Fr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   53
         Top             =   1365
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   52
         Top             =   1365
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3280
         TabIndex        =   51
         Top             =   520
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "WS No From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   49
         Top             =   795
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Recp No from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   48
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   47
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3300
         TabIndex        =   45
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Group Company"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Receipt Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   41
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period Mth Fr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   40
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   70
      Top             =   6360
      Width           =   11655
      Begin VB.TextBox txtActOPFee 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   74
         Top             =   345
         Width           =   1215
      End
      Begin VB.TextBox txtActOPFeeDesc 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   73
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtActOPCLMDesc 
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
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   72
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtActOPCLM 
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
         Left            =   6525
         MaxLength       =   10
         TabIndex        =   71
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label Label19 
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
         Index           =   2
         Left            =   600
         TabIndex        =   80
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Description"
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
         Index           =   1
         Left            =   2040
         TabIndex        =   79
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "OP Fee"
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
         Index           =   7
         Left            =   120
         TabIndex        =   78
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "OP Claims"
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
         Index           =   8
         Left            =   5760
         TabIndex        =   77
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label Label19 
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
         Index           =   9
         Left            =   6600
         TabIndex        =   76
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Description"
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
         Index           =   10
         Left            =   7800
         TabIndex        =   75
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   32
      Top             =   7080
      Width           =   11655
      Begin VB.TextBox lblSelectedAmt 
         Height          =   285
         Left            =   3600
         TabIndex        =   82
         Top             =   225
         Width           =   1335
      End
      Begin VB.TextBox lblcurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   81
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "&Print to File"
         Height          =   330
         Left            =   9240
         TabIndex        =   29
         Top             =   220
         Width           =   1095
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   6600
         TabIndex        =   25
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7200
         TabIndex        =   26
         Top             =   220
         Width           =   735
      End
      Begin VB.TextBox TxtPStatus2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   30
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   31
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8640
         TabIndex        =   28
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         Height          =   330
         Left            =   7920
         TabIndex        =   27
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   6000
         TabIndex        =   24
         Top             =   220
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   2520
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   7770
      Width           =   1695
   End
End
Attribute VB_Name = "FrmReceiptPayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim BordxNoArray() As String
Dim one As Boolean
Dim two As Boolean
Dim Check1 As Boolean
Dim Check2 As Boolean
Dim ScanCode As String
Dim listloop1 As Integer
Private m_blnDirection(1 To 16) As Boolean





Private Sub ChkScan_Click()
    If OptBordx.Value = True And ChkScan.Value = 0 Then
        Label14.Visible = True
        TxtBordx2.Visible = True
        ChkScan.Left = 5400
        ChkScan.Top = 800
        TxtBordx1.SetFocus
    ElseIf OptBordx.Value = True And ChkScan.Value = 1 Then
        Label14.Visible = False
        TxtBordx2.Visible = False
        ChkScan.Left = 3360
        ChkScan.Top = 800
        TxtBordx1.SetFocus
    ElseIf OptWks.Value = True And ChkScan.Value = 0 Then
        Label21.Visible = True
        TxtWSNoTo.Visible = True
        txtWSNoFr.SetFocus
        ChkScan.Left = 5400
        ChkScan.Top = 530
    ElseIf OptWks.Value = True And ChkScan.Value = 1 Then
        Label21.Visible = False
        TxtWSNoTo.Visible = False
        ChkScan.Left = 3360
        ChkScan.Top = 530
    End If
End Sub

Private Sub cisdbcbopayor_LostFocus()
Dim Company As New ADODB.Recordset
cboinsuredclient.Clear
Company.Open "Select GRPCoName,ASOINSPaycode from CISDB..grpcompany where ASOINSPaycode='" & Trim(Mid(cisdbcbopayor, 1, IIf(InStr(1, cisdbcbopayor, "-") = 0, 4, InStr(1, cisdbcbopayor, "-") - 1))) & "' AND ASOINSPAYCODE<>'DC' ", MediConnCISDB, adOpenKeyset, adLockOptimistic
    If Len(Trim(cisdbcbopayor)) Then
        Do Until Company.EOF
            cboinsuredclient.AddItem Trim(Company!GrpCoName)
            Company.MoveNext
           Loop
           Company.Close
           Set Company = Nothing
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


Private Sub CmdEdit_Click()
    Screen.MousePointer = vbHourglass
    If OptBordx.Value = True Then
        one = True
    Else
        two = True
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdUpdate_Click()
Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs

    Screen.MousePointer = vbHourglass
    If OptBordx.Value = True Then
        Call UpdateBordx
    Else
        Call UpdateWorksheet
    End If
    Screen.MousePointer = vbDefault
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub Form_Load()
Call ActionLogUpdate(12, "FrmReceiptPayor")
    one = False
    two = False
    Call ComboListing
    Call PayorComboListing
    ReDim BordxNoArray(0)
    ReDim BordxNoArray(0)
    lstWorksheet.Visible = False
    '==============================
    Label14.Visible = False
    TxtBordx2.Visible = False
    ChkScan.Left = 3360
    ChkScan.Top = 800
    ChkScan.Enabled = True
    '==============================
    CboBordxStatus.ListIndex = 0
    CboWSStatus.Visible = False
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    lblSelectedAmt = 0
    Call CIS_PayorComboListing
End Sub
Private Sub CboPayor_LostFocus()
Const METHOD_NAME = "CboPayor_LostFocus"
On Error GoTo debugs

Dim sql As String
Dim sql2 As String
Dim GRPCOMName As String
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strordergroup As String
  
    sql = ""
    GRPCOMName = ""
    GRPCOMName = Replace(CboPayor, "'", "''")
    '=Run Sp=================================================
    sql2 = ""
    sql2 = "Exec SearchSubsidiary '" & Trim(GRPCOMName) & "'"
    rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    '===================================================
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

'************** Start Command Button ****************************
Private Sub cmdSearch_Click()
Const METHOD_NAME = "cmdSearch_Click"
On Error GoTo debugs
If Mid(TxtBordx1, 1, 2) = "LO" Or Mid(txtWSNoFr, 1, 2) = "LO" Then
    If Mid(TxtBordx1, 1, 2) = "LO" Then
        ScanCode = TxtBordx1
        TxtBordx1 = ""
    ElseIf Mid(txtWSNoFr, 1, 2) = "LO" Then
        ScanCode = txtWSNoFr
        txtWSNoFr = ""
    End If
Else
    one = False
    two = False
    CmdEdit.Enabled = False
    CmdUpdate.Enabled = False
    cmdEXPtoExcel.Enabled = False
    CmdSave.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    ReDim BordxNoArray(0)
    Screen.MousePointer = vbHourglass
        If ChkScan.Value = 1 And OptWks.Value = True Then
            If txtWSNoFr = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                txtWSNoFr.SetFocus
                CmdExit.Enabled = True
                CmdEdit.Enabled = True
                CmdUpdate.Enabled = True
                cmdEXPtoExcel.Enabled = True
                CmdSave.Enabled = True
                CmdClear.Enabled = True
            Else
                Screen.MousePointer = vbHourglass
                listrecord = lstWorksheet.ListItems.Count
                If listrecord <> 0 Then
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                If Check1 = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    CmdExit.Enabled = True
                    CmdEdit.Enabled = True
                    CmdUpdate.Enabled = True
                    cmdEXPtoExcel.Enabled = True
                    CmdSave.Enabled = True
                    CmdClear.Enabled = True
                ElseIf Check1 = False Then
                    Call SearchRecord
                End If
                Check1 = False
                Screen.MousePointer = Default
                txtWSNoFr = ""
                txtWSNoFr.SetFocus
            End If
            txtWSNoFr.SetFocus
        ElseIf ChkScan.Value = 1 And OptBordx.Value = True Then
            If TxtBordx1 = "" Then
                MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtBordx1.SetFocus
                CmdExit.Enabled = True
                CmdEdit.Enabled = True
                CmdUpdate.Enabled = True
                cmdEXPtoExcel.Enabled = True
                CmdSave.Enabled = True
                CmdClear.Enabled = True
            Else
                Screen.MousePointer = vbHourglass
                listrecord = lstBordx.ListItems.Count
                If listrecord <> 0 Then
                    For listloop1 = 1 To listrecord
                    Call BordxNoCheck
                    Next listloop1
                End If
                If Check2 = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    CmdExit.Enabled = True
                    CmdEdit.Enabled = True
                    CmdUpdate.Enabled = True
                    cmdEXPtoExcel.Enabled = True
                    CmdSave.Enabled = True
                    CmdClear.Enabled = True
                ElseIf Check2 = False Then
                    Call SearchRecord
                End If
                Check2 = False
                Screen.MousePointer = Default
                TxtBordx1 = ""
                TxtBordx1.SetFocus
            End If
            TxtBordx1.SetFocus
        Else
            Call SearchRecord
        End If
        Screen.MousePointer = Default
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
End Sub

Private Sub cmdClear_Click()
    lstBordx.ListItems.Clear
    lstWorksheet.ListItems.Clear
    ReDim BordxNoArray(0)
    CboPayor.Clear
    Call PayorComboListing
    cboAccPeriodMth.Text = ""
    cboAccPeriodYR.Text = ""
    TxtBordx1.Text = ""
    TxtBordx2.Text = ""
    txtReceiptNo.Text = ""
    txtReceiptDate.Text = "__/__/____"
    txtChequeNo.Text = ""
    TxtRemark.Text = ""
    CboBordxStatus.ListIndex = 0
    CboWSStatus.ListIndex = 0
    lblCurrent = 0
    lblTotal = 0
    txtReceiptAmt = ""
    lblSelectedAmt = 0
    one = False
    two = False
    CmdSave.Enabled = True
    ChkScan.Value = 1
    cisdbcbopayor.Clear
    cboInsured.ListIndex = 0
    Call CIS_PayorComboListing
    'Call ClearForm(Me)
    If CboPayor <> "" Then
        cboinsuredclient.ListIndex = 0
    End If
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Const METHOD_NAME = "CmdSave_Click"
On Error GoTo debugs
    Screen.MousePointer = vbHourglass
    If OptBordx.Value = True Then
        Call SaveBordx
    Else
        Call SaveWorksheet
    End If
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdEXPtoExcel_Click()
    Screen.MousePointer = vbHourglass
    If OptBordx.Value = True Then
        Call ExportToExcelReceiptPayorBor(lstBordx)
    Else
        Call ExportToExcelReceiptPayorWks(lstWorksheet)
    End If
    Screen.MousePointer = vbDefault
End Sub
'************** End Command Button ****************************

'************** Start SearchRecord ****************************
Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim StrSql As String
Dim strAppend As String
Dim sql As String
Dim GRPCompany As String

    strAppend = ""
    GRPCompany = ""
    
    GRPCompany = Replace(CboPayor, "'", "''")

    If OptBordx.Value = True Then
        If Len(Trim(GRPCompany)) Then
            strAppend = strAppend + " And Bordx.GroupCompany = '" & Trim(GRPCompany) & "'"
        End If
    Else
       If Len(Trim(GRPCompany)) Then
            strAppend = strAppend + " And Bordx.GroupCompany = '" & Trim(GRPCompany) & "'"
        End If
    End If
    
    If Len(Trim(cboSubsidiary)) Then
        strAppend = strAppend + " AND MBMOthers.Subsidiary = '" & Trim(cboSubsidiary) & "'"
    End If
    
    If Len(Trim(TxtBordx1)) Then
      strAppend = strAppend + " and Bordx.BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
      strAppend = strAppend + " and Bordx.BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
    End If
    
    If Mid(CboWSStatus, 1, 1) = 2 Then
      strAppend = strAppend + " and COALESCE(Claim.PVRecPayorStatus , '')='' "
    End If
    
    If Mid(CboWSStatus, 1, 1) = 1 Then
     strAppend = strAppend + " and Claim.PVRecPayorStatus = 'F'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 1 Then
     ''strAppend = strAppend + " and Bordx.BordxPAYFrStatus = 'F'"
     strAppend = strAppend + " and COALESCE(ReceiptfrPAY.ReceiptNo, '') <> ''"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 2 Then
     ''strAppend = strAppend + " and COALESCE(Bordx.BordxPAYFrStatus , '')='' "
       strAppend = strAppend + " and COALESCE(ReceiptfrPAY.ReceiptNo, '') = ''"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 3 Then
     strAppend = strAppend + " and Bordx.BordxPAYFrStatus = 'P'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 4 Then
     strAppend = strAppend + " and (Bordx.BordxPAYFrStatus = 'P' or Bordx.BordxPAYFrStatus = 'F') "
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 5 Then
     strAppend = strAppend + " and ((Bordx.BordxPAYFrStatus , '')='' or BordxPAYFrStatus = 'P') "
    End If
    
    If Len(Trim(txtRECNoFr)) Then
       strAppend = strAppend + " AND ReceiptfrPAY.ReceiptNo >= '" & RemStr(Trim(txtRECNoFr)) & "'"
    End If
       
    If Len(Trim(txtRECNoTo)) Then
       strAppend = strAppend + " AND ReceiptfrPAY.ReceiptNo <= '" & RemStr(Trim(txtRECNoTo)) & "'"
    End If
    
    
    If Len(Trim(txtRECDateFr)) > 0 And IsDate(txtRECDateFr) = True Then
        strAppend = strAppend + " And ReceiptfrPAY.ReceiptDate >= '" & Format(txtRECDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRECDateTo)) > 0 And IsDate(txtRECDateTo) = True Then
        strAppend = strAppend + " And ReceiptfrPAY.ReceiptDate <= '" & Format(txtRECDateTo, "mm/dd/yyyy") & "'"
    End If
       
    If Len(cboACPeriodMthFr) Then
       strAppend = strAppend + " AND ReceiptfrPAY.RECAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
       
    If Len(cboACPeriodMthTo) Then
       strAppend = strAppend + " AND ReceiptfrPAY.RECAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If Len(cboACPeriodYRFr) Then
       strAppend = strAppend + " AND ReceiptfrPAY.RECAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
       
    If Len(cboACPeriodYRTo) Then
       strAppend = strAppend + " AND ReceiptfrPAY.RECAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
            
    If ChkScan.Value = 1 And OptWks.Value = True Then
        If Len(txtWSNoFr) Then
           strAppend = strAppend + " AND CISCas.CASNumber = '" & Trim(txtWSNoFr) & "'"
        End If
        
        If Len(Trim(TxtBordx1)) Then
          strAppend = strAppend + " and Bordx.BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
          strAppend = strAppend + " and Bordx.BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
    ElseIf ChkScan.Value = 0 And OptWks.Value = True Then
        If Len(txtWSNoFr) Then
           strAppend = strAppend + " AND CISCas.CASNumber >= '" & Trim(txtWSNoFr) & "'"
        End If
        If Len(TxtWSNoTo) Then
           strAppend = strAppend + " AND CISCas.CASNumber <= '" & Trim(TxtWSNoTo) & "'"
        End If
        
        If Len(Trim(TxtBordx1)) Then
          strAppend = strAppend + " and Bordx.BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
          strAppend = strAppend + " and Bordx.BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
    End If
    
    If ChkScan.Value = 1 And OptBordx.Value = True Then
        If Len(Trim(TxtBordx1)) Then
          strAppend = strAppend + " and Bordx.BordxRefNo = '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(txtWSNoFr) Then
           strAppend = strAppend + " AND CISCas.CASNumber >= '" & Trim(txtWSNoFr) & "'"
        End If
        If Len(TxtWSNoTo) Then
           strAppend = strAppend + " AND CISCas.CASNumber <= '" & Trim(TxtWSNoTo) & "'"
        End If
        
    ElseIf ChkScan.Value = 0 And OptBordx.Value = True Then
        If Len(Trim(TxtBordx1)) Then
          strAppend = strAppend + " and Bordx.BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
          strAppend = strAppend + " and Bordx.BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
        If Len(txtWSNoFr) Then
           strAppend = strAppend + " AND CISCas.CASNumber >= '" & Trim(txtWSNoFr) & "'"
        End If
        If Len(TxtWSNoTo) Then
           strAppend = strAppend + " AND CISCas.CASNumber <= '" & Trim(TxtWSNoTo) & "'"
        End If
    End If
    
    'cisdbcbopayor
    If Len(Trim(cisdbcbopayor)) Then
        strAppend = strAppend + " And  GRPCompany.ASOINSPaycode = '" & Trim(Mid(cisdbcbopayor, 1, IIf(InStr(1, cisdbcbopayor, "-") = 0, 4, InStr(1, cisdbcbopayor, "-") - 1))) & "'"
    End If
    'cboinsuredclient
    If Len(Trim(cboinsuredclient)) Then
        strAppend = strAppend + " And Bordx.GroupCompany = '" & Trim(cboinsuredclient) & "'"
    End If
    ' New Enhancement
    If Len(ClientPVNofrom) Then
       strAppend = strAppend + " AND ReceiptfrPAY.ClientPVNo >= '" & Trim(ClientPVNofrom) & "'"
    End If
    
    If Len(ClientPVNoto) Then
       strAppend = strAppend + " AND ReceiptfrPAY.ClientPVNo <= '" & Trim(ClientPVNoto) & "'"
    End If
    
    If Len(MedixReceiptNofrom) Then
       strAppend = strAppend + " AND ReceiptfrPAY.MedixReceiptNo >= '" & Trim(MedixReceiptNofrom) & "'"
    End If
    
    If Len(MedixReceiptNoto) Then
       strAppend = strAppend + " AND ReceiptfrPAY.MedixReceiptNo <= '" & Trim(MedixReceiptNoto) & "'"
    End If
    
    If cboInsured = "YES" Then
         strAppend = strAppend + " AND substring(Bordx.BordxRefNo, 1, 3) = 'KDC' "
    ElseIf cboInsured = "NO" Then
        strAppend = strAppend + " AND substring(Bordx.BordxRefNo, 1, 3) <> 'KDC' "
    End If
    
    If OptBordx.Value = True Then
        Call ReceiptBordxListing(strAppend)
    ElseIf OptWks.Value = True Then
        Call ReceiptWorksheetListing(strAppend)
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************** End SearchRecord ****************************

'************** Start SaveBordx ****************************
Private Function SaveBordx()
Const METHOD_NAME = "SaveBordx"
On Error GoTo debugs

Dim ii As Integer
Dim update
Dim BordxNo As String
Dim CaseNo As String
Dim status As String
Dim delimiter1 As Integer
Dim strsqlBordx As String
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Bordx As New ADODB.Recordset
    
    If lstBordx.ListItems.Count = 0 And txtActOPCLM <> "" And txtActOPCLM <> "0" Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf CboPayor = "" And txtActOPCLM <> "" And txtActOPCLM <> "0" Then
        MsgBox "Please select group company! ", vbCritical
    ElseIf CboPayor = "" And txtActOPFee <> "" And txtActOPFee <> "0" Then
        MsgBox "Please select group company!", vbCritical
    ElseIf txtReceiptNo = "" Then
       MsgBox "Please Enter Receipt No!", vbCritical
       txtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
       MsgBox "Please Enter Receipt Date!", vbCritical
       txtReceiptDate.SetFocus
    ElseIf txtChequeNo = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
       txtChequeNo.SetFocus
    ElseIf txtReceiptAmt = "" Then
        MsgBox "Please Enter Receipt Amount!", vbCritical
        txtReceiptAmt.SetFocus
    ElseIf txtReceiptAmt <> lblSelectedAmt And TxtRemark = "" Then
        MsgBox "Please Enter Remark!", vbCritical
        TxtRemark.SetFocus
    Else
        update = MsgBox("Do you want to update the record ?", vbYesNo + vbExclamation)
        If update = vbYes Then
        
            For ii = 0 To UBound(BordxNoArray)
                
                If Trim(BordxNoArray(ii)) <> "" Then
                   delimiter1 = InStr(1, BordxNoArray(ii), "~")
                   CaseNo = Trim(Mid(BordxNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                   BordxNo = Trim(Mid(BordxNoArray(ii), delimiter1 + 1, Len(BordxNoArray(ii)) - delimiter1))
                   Call UpdateRecord(CaseNo, BordxNo)
                End If
            Next ii
                If txtActOPCLM = "" Then
                    txtActOPCLM = 0
                End If
                If Len(txtActOPCLM) And txtActOPCLM <> 0 Then
                    strsqlREC = "Insert Into ReceiptSummary(ReceiptNo, ReceiptDate, ReceiptAmt,ReceiptType,ReceiptDesc,RecChqNo,RecPayCode) " & _
                         "Values('" & Trim(txtReceiptNo) & "','" & Format(txtReceiptDate, "yyyy/mm/dd") & "'," & txtActOPCLM & ", 'OPCLM', '" & txtActOPCLMDesc & "', '" & txtChequeNo & "', '" & CboPayor & "')"
                    MediConnCISPayment.Execute strsqlREC
                End If
                If txtActOPFee = "" Then
                    txtActOPFee = 0
                End If
                If Len(txtActOPFee) And txtActOPFee <> 0 Then
                    strsqlREC = "Insert Into ReceiptSummary(ReceiptNo, ReceiptDate, ReceiptAmt,ReceiptType,ReceiptDesc,RecChqNo,RecPayCode) " & _
                         "Values('" & Trim(txtReceiptNo) & "','" & Format(txtReceiptDate, "yyyy/mm/dd") & "'," & txtActOPFee & ", 'OPFEE', '" & txtActOPFeeDesc & "', '" & txtChequeNo & "', '" & CboPayor & "')"
                    MediConnCISPayment.Execute strsqlREC
                End If
                
            Erase BordxNoArray
            Arraycnt = 0
            ReDim BordxNoArray(0)
            lstBordx.ListItems.Clear
            Call ReceiptBordxListing(" AND ReceiptfrPAY.ReceiptNo = '" & txtReceiptNo & "'")
            txtReceiptNo.Text = ""
            txtReceiptDate.Text = "__/__/____"
            txtChequeNo.Text = ""
            TxtRemark.Text = ""
            txtReceiptAmt = ""
            cboAccPeriodMth.Text = ""
            cboAccPeriodYR.Text = ""
            lblSelectedAmt = 0
            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************** End SaveBordx ****************************

Private Function UpdateBordx()
Const METHOD_NAME = "UpdateBordx"
On Error GoTo debugs
Dim RecordSaved
Dim RptPAY As New ADODB.Recordset
Dim UpdateRptPAY As New ADODB.Recordset
Dim RptBOR As New ADODB.Recordset
Dim UpdateRptBOR As New ADODB.Recordset
Dim StrSql As String
Dim StrSql2 As String
Dim Index As String
Dim Bordx As String
Dim subsidary As String
    If one = False Then
        ElseIf lstBordx.ListItems.Count = 0 Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        ElseIf txtReceiptNo = "" Then
            MsgBox "Please Enter Receipt No!", vbCritical
            txtReceiptNo.SetFocus
        ElseIf IsDate(txtReceiptDate) = False Then
            MsgBox "Invalid Receipt Date!", vbCritical
            txtReceiptDate.SetFocus
        ElseIf txtReceiptDate = "__/__/____" Then
            MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtReceiptDate.SetFocus
        ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque No!", vbCritical
            txtChequeNo.SetFocus
        ElseIf txtReceiptAmt = "" Then
            MsgBox "Please Enter Receipt Amt!", vbCritical
            txtReceiptAmt.SetFocus
        Else
            
            Bordx = Trim(lstBordx.SelectedItem.SubItems(2))
            Index = Trim(lstBordx.SelectedItem.SubItems(4))
            subsidary = Trim(lstBordx.SelectedItem.SubItems(8))
            
            StrSql = "Select * From ReceiptFrPAY Where BordxNo = '" & Trim(Bordx) & "'" 'AND ReceiptNo = '" & Trim(Index) & "'"
            StrSql2 = "Select * From Bordx Where BordxRefNo = '" & Trim(Bordx) & "'" '' AND ReceiptNo = '" & Trim(Index) & "'"
                                 
            RptPAY.Open StrSql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            RptBOR.Open StrSql2, MediConnCISWS, adOpenKeyset, adLockOptimistic
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                UpdateRptPAY.Open "ReceiptFrPAY", MediConnCISPayment, adOpenKeyset, adLockOptimistic
                If RptPAY.RecordCount = 0 Then
                    RptPAY.AddNew
                End If
                Do Until RptPAY.EOF
                    With RptPAY
                        !ReceiptNo = ChkStr(txtReceiptNo)
                        !ReceiptDate = txtReceiptDate
                        !ReceiptAmt = txtReceiptAmt
                        !ChequeNo = ChkStr(txtChequeNo)
                        !Remark = ChkStr(TxtRemark)
                        !PAYLastUpdateDate = Date
                        !PAYLastUpdateUser = Logon
                        If cboAccPeriodMth <> "" Then
                            !RECAccPeriodMth = cboAccPeriodMth
                        End If
                        If cboAccPeriodYR <> "" Then
                            !RECAccPeriodYR = cboAccPeriodYR
                        End If
                        .update
                    End With
                    RptPAY.MoveNext
                Loop
                
                With RptBOR
                    !BordxPAYRecDate = txtReceiptDate
                End With
                
                MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
          
                RptPAY.Close
                Set UpdateRptPAY = Nothing
                'RptBOR.Close
                Set UpdateRptBOR = Nothing
                
                txtReceiptNo.Text = ""
                txtReceiptDate.Text = "__/__/____"
                txtChequeNo.Text = ""
                TxtRemark.Text = ""
                txtReceiptAmt = ""
                cboAccPeriodMth.Text = ""
                cboAccPeriodYR.Text = ""
                lblSelectedAmt = 0
                Call SearchRecord
                one = False
                CmdSave.Enabled = True
            End If
            Set RptPAY = Nothing
            CmdSave.Enabled = True
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

'************** Start SaveWorksheet ****************************

Private Function SaveWorksheet()
Const METHOD_NAME = "SaveWorksheet"
On Error GoTo debugs

Dim casno As String
Dim CASVer As String
Dim status As String
Dim BordxNo As String
Dim PAYCode As String
Dim ii As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim delimiter1, delimiter2, delimiter3 As Integer
Dim ClmCheck1 As New ADODB.Recordset
Dim ClmCheck2 As New ADODB.Recordset
Dim update
Dim intloop1 As Integer
Dim intloop2 As Integer
    
    
    If lstWorksheet.ListItems.Count = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf txtReceiptNo = "" Then
       MsgBox "Please Enter Receipt No !", vbCritical
       txtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf txtChequeNo = "" Then
       MsgBox "Please Enter Cheque No. !", vbCritical
       txtChequeNo.SetFocus
    ElseIf txtReceiptAmt <> lblSelectedAmt And TxtRemark = "" Then
        MsgBox "Please Enter Remark!", vbCritical
        TxtRemark.SetFocus
    Else
        update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If update = vbYes Then
                For ii = 0 To UBound(BordxNoArray)
                   If Trim(BordxNoArray(ii)) <> "" Then
                        delimiter1 = InStr(1, BordxNoArray(ii), "~")
                        casno = Trim(Mid(BordxNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                        delimiter2 = InStr(delimiter1 + 1, BordxNoArray(ii), "~")
                        BordxNo = Trim(Mid(BordxNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                        delimiter3 = InStr(delimiter2 + 1, BordxNoArray(ii), "~")
                        status = Trim(Mid(BordxNoArray(ii), delimiter2 + 1, delimiter3 - delimiter2 - 1))
                        PaymentAmt = Trim(Mid(BordxNoArray(ii), delimiter3 + 1, delimiter2))
                        Call UpdateRecord2(casno, BordxNo, status, PaymentAmt)
                        STRCHECK1 = " Select COUNT(*) AS CNtTot " & _
                                   " From Claim Where BordxRefno = '" & Trim(BordxNo) & "'"
                        
                        STRCHECK2 = " Select COUNT(*) AS CNtTot From Claim " & _
                                   " Where PVRecPayorStatus = 'F' and BordxRefno = '" & Trim(BordxNo) & "'"
                        
                        Set ClmCheck1 = MediConnCISWS.Execute(STRCHECK1)
                        Set ClmCheck2 = MediConnCISWS.Execute(STRCHECK2)
                        intloop1 = 0
                        intloop2 = 0
                        Do While Not ClmCheck1.EOF
                            intloop1 = ClmCheck1!CNtTot
                            ClmCheck1.MoveNext
                        Loop
                        ClmCheck1.Close
                        Do While Not ClmCheck2.EOF
                           intloop2 = ClmCheck2!CNtTot
                           ClmCheck2.MoveNext
                        Loop
                        ClmCheck2.Close
                        If intloop1 = intloop2 Then
                           Call UpdateBordxStatus(BordxNo, "F")
                        Else
                            Call UpdateBordxStatus(BordxNo, "P")
                        End If
                    End If
               Next ii
           
                'If Len(txtActOPCLM) And txtActOPCLM <> 0 Then
                 '   strsqlREC = "Insert Into ReceiptSummary(ReceiptNo, ReceiptDate, ReceiptAmt,ReceiptType,ReceiptDesc,RecChqNo,RecPayCode) " & _
                  '       "Values('" & Trim(TxtReceiptNo) & "','" & Format(txtReceiptDate, "yyyy/mm/dd") & "'," & txtActOPCLM & ", 'OPCLM', '" & txtActOPCLMDesc & "', '" & txtChequeNo & "', '" & CboPayor & "')"
                   ' MediConnCISPayment.Execute strsqlREC
                'End If
                'If Len(txtActOPFee) And txtActOPFee <> 0 Then
                 '   strsqlREC = "Insert Into ReceiptSummary(ReceiptNo, ReceiptDate, ReceiptAmt,ReceiptType,ReceiptDesc,RecChqNo,RecPayCode) " & _
                  '       "Values('" & Trim(TxtReceiptNo) & "','" & Format(txtReceiptDate, "yyyy/mm/dd") & "'," & txtActOPFee & ", 'OPFEE', '" & txtActOPFeeDesc & "', '" & txtChequeNo & "', '" & CboPayor & "')"
                   ' MediConnCISPayment.Execute strsqlREC
                'End If
           Erase BordxNoArray
           Arraycnt = 0
           ReDim BordxNoArray(0)
           lstWorksheet.ListItems.Clear
           Call ReceiptWorksheetListing(" AND ReceiptfrPAY.ReceiptNo = '" & txtReceiptNo & "'")
           txtReceiptNo.Text = ""
           txtReceiptDate.Text = "__/__/____"
           txtChequeNo.Text = ""
           TxtRemark.Text = ""
           txtReceiptAmt = ""
           cboAccPeriodMth.Text = ""
           cboAccPeriodYR.Text = ""
           lblSelectedAmt = 0
            MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************** End SaveWorksheet ****************************

Private Function UpdateWorksheet()
Const METHOD_NAME = "UpdateWorksheet"
On Error GoTo debugs

Dim RecordSaved
Dim WKs As New ADODB.Recordset
Dim UpdateWKs As New ADODB.Recordset
Dim BOR As New ADODB.Recordset
Dim UpdateBOR As New ADODB.Recordset
Dim StrSql As String
Dim StrSql2 As String
Dim casno As String
Dim WSNo As String
    
    If two = False Then
        ElseIf lstWorksheet.ListItems.Count = 0 Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        ElseIf txtReceiptNo = "" Then
            MsgBox "Please Enter Receipt No!", vbCritical
            txtReceiptNo.SetFocus
        ElseIf IsDate(txtReceiptDate) = False Then
            MsgBox "Invalid Receipt Date!", vbCritical
            txtReceiptDate.SetFocus
        ElseIf txtReceiptDate = "__/__/____" Then
            MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtReceiptDate.SetFocus
        ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque No!", vbCritical
            txtChequeNo.SetFocus
        ElseIf txtReceiptAmt = "" Then
            MsgBox "Please Enter Receipt Amt!", vbCritical
            txtReceiptAmt.SetFocus
        Else
            Bordx = Trim(lstWorksheet.SelectedItem.SubItems(2))
            casno = Trim(lstWorksheet.SelectedItem.SubItems(3))
            StrSql = " Select * From ReceiptFrPAY Where BordxNo = '" & Trim(Bordx) & "' AND CASNumber = '" & Trim(casno) & "'"
            StrSql2 = " Select * From Bordx Where BordxRefNo = '" & Trim(Bordx) & "'"
            WKs.Open StrSql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            BOR.Open StrSql2, MediConnCISWS, adOpenKeyset, adLockOptimistic
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                UpdateWKs.Open "ReceiptFrPAY", MediConnCISPayment, adOpenKeyset, adLockOptimistic
                If WKs.RecordCount = 0 Then
                    WKs.AddNew
                End If
                Do Until WKs.EOF
                    With WKs
                        !ReceiptNo = ChkStr(txtReceiptNo)
                        !ReceiptDate = txtReceiptDate
                        !ReceiptAmt = txtReceiptAmt
                        !ChequeNo = ChkStr(txtChequeNo)
                        !Remark = ChkStr(TxtRemark)
                        !PAYLastUpdateDate = Date
                        !PAYLastUpdateUser = Logon
                        .update
                    End With
                    WKs.MoveNext
                Loop
                With BOR
                    !BordxPAYRecDate = txtReceiptDate
                End With
                MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
                WKs.Close
                Set UpdateWKs = Nothing
                Set UpdateBOR = Nothing
                txtReceiptNo.Text = ""
                txtReceiptDate.Text = "__/__/____"
                txtChequeNo.Text = ""
                TxtRemark.Text = ""
                txtReceiptAmt = ""
                cboAccPeriodMth.Text = ""
                cboAccPeriodYR.Text = ""
                lblSelectedAmt = 0
                Call SearchRecord
                two = False
                CmdSave.Enabled = True
            End If
            Set WKs = Nothing
            
        CmdSave.Enabled = True
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
Private Sub lstBordx_Click()
Const METHOD_NAME = "lstBordx_Click"
On Error GoTo debugs
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
Dim BordxNo As String
Dim ReceiptNo As String
If one = True Then
CmdSave.Enabled = False
    If lstBordx.ListItems.Count Then
        BordxNo = Trim(lstBordx.SelectedItem.SubItems(2))
        StrSql = " SELECT * From ReceiptFrPAY Where BordxNo = '" & Trim(BordxNo) & "'"
        rst3.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockReadOnly
            If rst3.EOF = False Then
                MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                txtReceiptNo.SetFocus
                txtReceiptNo = rst3!ReceiptNo
                txtReceiptDate = rst3!ReceiptDate
                txtReceiptAmt = rst3!ReceiptAmt
                txtChequeNo = rst3!ChequeNo
                TxtRemark = rst3!Remark
                If Len(rst3!RECAccPeriodMth) Then
                    cboAccPeriodMth = rst3!RECAccPeriodMth
                End If
                If Len(rst3!RECAccPeriodYR) Then
                    cboAccPeriodYR = rst3!RECAccPeriodYR
                End If
        
                CmdSave.Enabled = False
                rst3.Close
                Set rst3 = Nothing
            End If
    End If
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'************** start ItemCheck ****************************

Private Sub lstBordx_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "lstBordx_ItemCheck"
On Error GoTo debugs

Dim i, newbound As Integer
Dim tempArray() As String
Dim Found As String
Dim ListLoop As Integer
Dim ListCount As Integer
Dim TotalAmt As Variant
Dim RECStatus As String
Dim RECAmt As Integer
Dim strsqlBOR As String
Dim strsqlREC As String
Dim BOR As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim BordxNo As String
Dim CaseNo As String
Dim CASVer As String
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Static Arraycnt As Integer

    
    If Item.Checked = True Then
        BordxNo = Trim(Mid(Item.SubItems(2), 1, IIf(InStr(1, Item.SubItems(2), "-") = 0, 10, InStr(1, Item.SubItems(2), "-") - 1)))
        strsqlBOR = "select BordxRefNo, BordxPAYFrStatus from Bordx where BordxRefNo = '" & Trim(BordxNo) & "'"
        BOR.Open strsqlBOR, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
        If BOR!BordxPAYFRStatus = "F" Then
            MsgBox "Record has been paid !", vbCritical
            Item.Checked = False
        Else
        
            If BOR!BordxPAYFRStatus = "P" Then
                MsgBox "Record has been Partial Paid !", vbInformation
                
                strsqlREC = "select * from ReceiptfrPAY where BordxNo = '" & Trim(BordxNo) & "'"
                REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockOptimistic
                
                strsqlCLM = " Select CASNumber from Claim where BordxRefNo = '" & Trim(BordxNo) & "'" & _
                            " AND PVRecPayorStatus IS NULL "
                CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
            
                If CLM.RecordCount Then
                    Do Until CLM.EOF
                        CaseNo = CLM!CasNumber
                                            
                        ReDim Preserve BordxNoArray(Arraycnt)
                        BordxNoArray(Arraycnt) = CLM!CasNumber & "~" & Item.SubItems(2)
                        Arraycnt = Arraycnt + 1
                    
                    CLM.MoveNext
                    Loop
                
                    If Item.SubItems(6) = "Full" Then
                    lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(5), "#,##0.00"))
                    Item.SubItems(6) = ""
                    End If
                
                    lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(5), "#,##0.00"))
                    
                    REC.Close
                    Set REC = Nothing
                    CLM.Close
                    Set CLM = Nothing
                End If
            
            Else
            
            strsqlREC = "select * from ReceiptfrPAY where BordxNo = '" & Trim(BordxNo) & "'"
            REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            
            'strsqlCLM = "Select CASNumber from Claim where BordxRefNo = '" & Trim(BordxNo) & "'"
            strsqlCLM = "Select CASNumber,BTBGBordxRefNo from Claim where BordxRefNo = '" & Trim(BordxNo) & "' or BTBGBordxRefNo = '" & Trim(BordxNo) & "'"
            CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
            If CLM.RecordCount Then
                Do Until CLM.EOF
                    CaseNo = CLM!CasNumber
                    ReDim Preserve BordxNoArray(Arraycnt)
                    BordxNoArray(Arraycnt) = Trim(CLM!CasNumber) & "~" & Item.SubItems(2)
                    Arraycnt = Arraycnt + 1
                CLM.MoveNext
                Loop
                If Item.SubItems(6) = "Full" Then
                lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(3), "#,##0.00"))
                Item.SubItems(6) = ""
                End If
            
                REC.Close
                Set REC = Nothing
                CLM.Close
                Set CLM = Nothing
            End If
            lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(5), "#,##0.00"))
            End If
        End If
        BOR.Close
        Set BOR = Nothing
    Else
        newbound = 0
        If UBound(BordxNoArray) <> 0 Then
            For i = 0 To UBound(BordxNoArray)
                If (Item.SubItems(2)) <> BordxNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = BordxNoArray(i)
                    newbound = newbound + 1
                End If
            Next
          
            If Item.SubItems(6) = "Full" Then
                lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(3), "#,##0.00"))
                Item.SubItems(6) = ""
            End If
          
            lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(3), "#,##0.00"))
            
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            ReDim tempArray(0)
          
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
          
            ReDim BordxNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                BordxNoArray(i) = tempArray(i)
            Next
            Erase tempArray
            ReDim tempArray(0)
        
        Else
            If Item.SubItems(6) = "Full" Then
                lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(3), "#,##0.00"))
                Item.SubItems(6) = ""
            Else
                Item.SubItems(6) = ""
            End If
            
            lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(3), "#,##0.00"))
          
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0

        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub lstWorksheet_Click()
Const METHOD_NAME = "lstWorksheet_Click"
On Error GoTo debugs

Dim rst3 As New ADODB.Recordset
Dim StrSql As String
Dim Bordx As String
Dim casno As String
Dim WSNo As String
If two = True Then
CmdSave.Enabled = False
    If lstWorksheet.ListItems.Count Then
        Bordx = Trim(lstWorksheet.SelectedItem.SubItems(2))
        casno = Trim(lstWorksheet.SelectedItem.SubItems(3))
        StrSql = " SELECT * From ReceiptFrPAY Where BordxNo = '" & Trim(Bordx) & "'" & _
                 " AND CASNumber = '" & Trim(casno) & "'"
        rst3.Open StrSql, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            If rst3.RecordCount = 0 Then
                MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                txtReceiptNo.SetFocus
                txtReceiptNo = rst3!ReceiptNo
                txtReceiptDate = rst3!ReceiptDate
                txtReceiptAmt = rst3!ReceiptAmt
                txtChequeNo = rst3!ChequeNo
                TxtRemark = rst3!Remark
                cboAccPeriodMth = rst3!RECAccPeriodMth
                cboAccPeriodYR = rst3!RECAccPeriodYR
                CmdSave.Enabled = False
                rst3.Close
                Set rst3 = Nothing
            End If
    End If
End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstWorksheet_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
Const METHOD_NAME = "ProcessCheckedItem"
On Error GoTo debugs

Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Dim ListLoop As Integer
Dim ListCount As Integer
Dim TotalAmt As Variant
Dim RECStatus As String
Dim RECAmt As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim strsqlREC As String
Dim REC As New ADODB.Recordset
Dim CaseNo As String
Dim caseversion As String
Dim status As String
Static Arraycnt As Integer
Dim BordxNo As String
Dim strsqlBOR As String
Dim BOR As New ADODB.Recordset
                        
    If Item.Checked = True Then
        BordxNo = Trim(Mid(Item.SubItems(2), 1, IIf(InStr(1, Item.SubItems(2), "-") = 0, 10, InStr(1, Item.SubItems(2), "-") - 1)))
        CaseNo = Trim(Item.SubItems(3))
        strsqlBOR = " Select BordxRefNo, PVRecPayorStatus from Claim where CASNumber = '" & Trim(CaseNo) & "'" & _
                    " AND BordxRefNo = '" & Trim(BordxNo) & "'"
        BOR.Open strsqlBOR, MediConnCISWS, adOpenKeyset, adLockOptimistic
        If BOR!PVRecPayorStatus = "F" Then
            MsgBox "Record has been paid !", vbCritical
            Item.Checked = False
        Else
            strsqlREC = "select * from ReceiptfrPAY where BordxNo = '" & Trim(BordxNo) & "'"
            REC.Open strsqlREC, MediConnCISPayment, adOpenKeyset, adLockOptimistic
            strsqlCLM = "Select CASNumber from Claim where BordxRefNo = '" & Trim(BordxNo) & "'"
            CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
                    If Item.SubItems(9) = "Partial" Then
                        lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Format(Item.SubItems(7), "#,##0.00"))
                    Else
                    Item.SubItems(9) = "Full"
                    lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Format(Item.SubItems(4), "#,##0.00"))
                    End If
                    ReDim Preserve BordxNoArray(Arraycnt)
                    BordxNoArray(Arraycnt) = Trim(CaseNo) & "~" & Trim(BordxNo) & "~" & Mid(Item.SubItems(9), 1, 1) & "~" & Item.SubItems(4) & Item.SubItems(5) & Item.SubItems(6)
                    Arraycnt = Arraycnt + 1
                REC.Close
                Set REC = Nothing
                CLM.Close
                Set CLM = Nothing
        End If
        BOR.Close
        Set BOR = Nothing
    Else
        newbound = 0
        If UBound(BordxNoArray) <> 0 Then
            For i = 0 To UBound(BordxNoArray)
                If (Item.SubItems(2)) <> BordxNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = BordxNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            If Item.SubItems(9) = "Full" Then
                lblSelectedAmt = CDbl(lblSelectedAmt) - CDbl(Format(Item.SubItems(4), "#,##0.00"))
                Item.SubItems(9) = ""
            End If
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            ReDim tempArray(0)
          
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
          
            ReDim BordxNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                BordxNoArray(i) = tempArray(i)
            Next
            Erase tempArray
            ReDim tempArray(0)
        
        Else
            If Item.SubItems(9) = "Full" Then
                lblSelectedAmt = CDbl(Format(lblSelectedAmt, "#,##0.00")) - CDbl(Format(Item.SubItems(4), "#,##0.00"))
                lblSelectedAmt = Format(lblSelectedAmt, "#,##0.00")
                Item.SubItems(9) = ""
            Else
                Item.SubItems(4) = ""
            End If
          
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0

        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function
'************** End ItemCheck ****************************

'************** Start ReceiptBordxListing ****************************
Private Sub ReceiptBordxListing(strAppend As String)
Const METHOD_NAME = "ReceiptBordxListing"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim total As String
Dim Current As String
Dim Check As Boolean
Dim status As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String

total = 0
Current = 0
lblSelectedAmt = 0
If ChkScan.Value = 0 Then
    lstBordx.ListItems.Clear
End If

strAppendCase = "18"
sql = sql + strAppend
Set cmd.ActiveConnection = MediConnCISPayment
cmd.CommandText = "PaymentReceiptModule"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 3000
Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
cmd.Parameters.Append Prm1
Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
cmd.Parameters.Append prm2
rst2.CursorLocation = adUseClient
rst2.CursorType = adOpenForwardOnly
rst2.CacheSize = 100
Set rst2 = cmd.Execute
Do
    Do Until rst2.EOF
    txtPStatus = ""
    If Len(rst2!BordxPAYFRStatus) And Len(rst2!ReceiptNo) Then
        txtPStatus = rst2!BordxPAYFRStatus
    Else
        txtPStatus = ""
    End If
    With rst2
        Set Record = lstBordx.ListItems.Add(, , "")
        If Len(!BordxPAYCode) Then
           Record.SubItems(1) = !BordxPAYCode
        End If
        If Len(!Bordxrefno) Then
           Record.SubItems(2) = Trim(!Bordxrefno)
        End If
        If Len(!BordxClaimAmt) Then
           Record.SubItems(3) = Format(!BordxClaimAmt, "#,##0.00")
        End If
        
        If Len(!CheckAmt) Then
           Record.SubItems(4) = Format(!CheckAmt, "#,##0.00")
        End If
        
        If Len(!CheckAmt) Then
            Record.SubItems(5) = Format(!BordxClaimAmt, "#,##0.00") - Format(!CheckAmt, "#,##0.00")
            Record.SubItems(5) = Format(Record.SubItems(5), "#,##0.00")
        Else
            Record.SubItems(5) = Format(!BordxClaimAmt, "#,##0.00")
        End If
        
        If txtPStatus = "F" Then
            Record.SubItems(6) = "Full"
        ElseIf txtPStatus = "P" Then
            Record.SubItems(6) = "Partial"
        ElseIf txtPStatus = "" Then
            Record.SubItems(6) = ""
        End If
        
        If Len(!BordxDate) Then
           Record.SubItems(7) = Format(!BordxDate, "dd/mm/yyyy")
        End If
        If Len(!Subsidiary) Then
           Record.SubItems(8) = Trim(!Subsidiary)
        End If
        If Len(!GroupCompany) Then
           Record.SubItems(9) = Trim(!GroupCompany)
        End If
        If Len(!ReceiptNo) Then
           Record.SubItems(10) = Trim(!ReceiptNo)
        End If
        
        If Len(!ClientPVNo) Then
           Record.SubItems(11) = Trim(!ClientPVNo)
        End If
        
        If Len(!MedixReceiptNo) Then
           Record.SubItems(12) = Trim(!MedixReceiptNo)
        End If
        
        
           Record.SubItems(13) = !isInsured
        
        
        Current = Current + 1
        lblCurrent = Current
    End With
    rst2.MoveNext
    DoEvents
        If intExit = 1 Then
            Check = False
            Exit Do
        End If
    Loop
Loop Until Check = False
intExit = 0
If lstBordx.ListItems.Count = 0 Then
    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    lblCurrent = 0
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
End If
rst2.Close
Set rst2 = Nothing
CmdExit.Enabled = True
CmdEdit.Enabled = True
CmdUpdate.Enabled = True
cmdEXPtoExcel.Enabled = True
CmdSave.Enabled = True
CmdClear.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'************** Start ReceiptWorksheetListing ****************************

Private Sub ReceiptWorksheetListing(strAppend As String)
Const METHOD_NAME = "ReceiptWorksheetListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim total As Integer
Dim Current As Integer
Dim Check As Boolean
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim strAppendCase As String
If ChkScan.Value = 0 Then
    lstWorksheet.ListItems.Clear
End If
strAppendCase = "19" ''3
sql = sql + strAppend
Set cmd.ActiveConnection = MediConnCISPayment
cmd.CommandText = "PaymentReceiptModule"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 3000
Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
cmd.Parameters.Append Prm1
Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
cmd.Parameters.Append prm2
rst2.CursorLocation = adUseClient
rst2.CursorType = adOpenForwardOnly
rst2.CacheSize = 100
Set rst2 = cmd.Execute
If rst2.EOF = True Then
    MsgBox "No record found !", vbOKOnly + vbCritical, DialogBoxTitle
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
End If
Do
    Do Until rst2.EOF
    txtPStatus = ""
    If Len(rst2!PVRecPayorStatus) Then
        txtPStatus = rst2!PVRecPayorStatus
    Else
        txtPStatus = ""
    End If
    With rst2
        Set Record = lstWorksheet.ListItems.Add(, , "")
               
        If Len(!BordxPAYCode) Then
           Record.SubItems(1) = !BordxPAYCode
        End If
        If Len(!Bordxrefno) Then
           Record.SubItems(2) = !Bordxrefno
        End If
        If Len(!CasNumber) Then
           Record.SubItems(3) = Trim(!CasNumber)
        End If
        'If TxtPStatus = "F" Then
         '   Record.SubItems(4) = Format(!CheckAmt, "#,##0.00")
        'Else
         '   Record.SubItems(4) = Format(!CLMBordxAmt, "#,##0.00")
        'End If
        Record.SubItems(4) = Format(!CLMBordxAmt, "#,##0.00")
        Record.SubItems(5) = Format(!CheckAmt, "#,##0.00")
        Record.SubItems(6) = Format(Format(!CLMBordxAmt, "#,##0.00") - IIf(IsNumeric(!CheckAmt), !CheckAmt, "0"), "#,##0.00")
        
        If Len(!ReceiptNo) Then
           Record.SubItems(7) = Trim(!ReceiptNo)
        End If
        If Len(!ReceiptDate) Then
           Record.SubItems(8) = !ReceiptDate
        End If
        If txtPStatus = "F" Then
            Record.SubItems(9) = "Full"
        ElseIf txtPStatus = "P" Then
            Record.SubItems(9) = "Partial"
        Else
            Record.SubItems(9) = ""
        End If
        If Len(!BordxDate) Then
           Record.SubItems(10) = Format(!BordxDate, "dd/mm/yyyy")
        End If
        If Len(!GroupCompany) Then
           Record.SubItems(11) = !GroupCompany
        End If
         If Len(!Subsidiary) Then
           Record.SubItems(12) = !Subsidiary
        End If
        
        If Len(!ClientPVNo) Then
           Record.SubItems(13) = Trim(!ClientPVNo)
        End If
        
        If Len(!MedixReceiptNo) Then
           Record.SubItems(14) = Trim(!MedixReceiptNo)
        End If
        
        Record.SubItems(15) = !isInsured
        
        Current = Current + 1
        lblCurrent = Current
    End With
    rst2.MoveNext
    DoEvents
        If intExit = 1 Then
            Check = False
            Exit Do
        End If
    Loop
Loop Until Check = False
intExit = 0
rst2.Close
Set rst2 = Nothing
CmdExit.Enabled = True
CmdEdit.Enabled = True
CmdUpdate.Enabled = True
cmdEXPtoExcel.Enabled = True
CmdSave.Enabled = True
CmdClear.Enabled = True

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub

Private Sub TxtBordx1_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtBordx1_LostFocus()
    TxtBordx1 = ChkStr(TxtBordx1)
    CmdSearch.Default = False
End Sub

Private Sub TxtBordx2_LostFocus()
TxtBordx2 = ChkStr(TxtBordx2)
End Sub

Private Sub txtChequeNo_LostFocus()
    txtChequeNo = ChkStr(txtChequeNo)
End Sub

'*************************** Start Diable Non Integer Value
Private Sub TxtReceiptAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

'*************************** Start Combobox Change/ KeyPress

Private Sub CboPayor_Change()
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Const METHOD_NAME = "CboPayor_KeyPress"
On Error GoTo debugs
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Dim StrSql As String
  If KeyAscii >= 32 And KeyAscii <> 127 Then
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboPayor.ListCount - 1
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
    Next i
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
'*************************** End Combobox Change/ KeyPress

'*************************** Start UpdateRecord

Private Sub UpdateRecord(CaseNo As String, BordxNo As String)
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs
Dim update
Dim strsqlBordx As String
Dim PAYAmt As String
Dim SearchDB As New ADODB.Recordset
Dim sql As String
Dim REceiptSql As String
    PAYAmt = 0
      sql = " Select  PVRecPayorStatus, case when BTBGBordxRefNo is not null then CLMBordxAmt - CLMBTBGAmt else CLMBordxAmt end as CLMBordxAmt  " & _
          " From Claim where CASNumber = '" & Trim(CaseNo) & "'" & _
          " And BordxRefNo = '" & Trim(BordxNo) & "' or BTBGBordxRefNo = '" & Trim(BordxNo) & "'"
    
    SearchDB.Open sql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    With SearchDB
        PAYAmt = IIf(IsNumeric(SearchDB!CLMBordxAmt), SearchDB!CLMBordxAmt, 0)
        !PVRecPayorStatus = "F"
    End With
    SearchDB.update
    Set SearchDB = Nothing
    
    REceiptSql = "Insert into ReceiptfrPAY(ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo," & _
                 "BordxNo, CASNumber, Remark, PAYLastUpdateDate, PAYLastUpdateUser, RECAccPeriodMth, " & _
                 "RECAccPeriodYR, RecPayAmt, CheckAmt) " & _
                 "Values('" & Trim(txtReceiptNo) & "','" & Format(txtReceiptDate, "yyyy/mm/dd") & _
                 "'," & txtReceiptAmt & ",'" & Trim(txtChequeNo) & "','" & BordxNo & _
                 "','" & CaseNo & "','" & Trim(TxtRemark) & "','" & Format(Date, "yyyy/mm/dd") & _
                 "','" & Logon & "','" & cboAccPeriodMth & "','" & cboAccPeriodYR & _
                 "'," & PAYAmt & "," & PAYAmt & ")"
    MediConnCISPayment.Execute REceiptSql
    Call UpdateBordxStatus(BordxNo, "F")
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateRecord2(casno As String, BordxNo As String, status As String, PaymentAmt As Variant)
Const METHOD_NAME = "UpdateRecord2"
On Error GoTo debugs
Dim update
Dim StrSql As String
Dim REC1 As New ADODB.Recordset
Dim REC2 As New ADODB.Recordset
Dim InvCheck As New ADODB.Recordset
Dim strsqlREC1 As String
Dim strsqlREC2 As String
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim tmpPayAmt As String

    strsqlREC1 = " Select * from ReceiptfrPAY " & _
                 " Where CASNumber = '" & Trim(casno) & "' AND BordxNo = '" & Trim(BordxNo) & "'"
    
    REC1.Open strsqlREC1, MediConnCISPayment, adOpenKeyset, adLockOptimistic
    
    strsqlCLM = " Select CASNumber, BordxRefNo, PVRecPayorStatus, CLMBordxAmt " & _
                " From Claim " & _
                " Where BordxRefNo = '" & Trim(BordxNo) & "' And CASNumber = '" & Trim(casno) & "'"
    CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
    tmpPayAmt = 0
    If REC1.RecordCount = 0 Then
        REC1.AddNew
    End If
        With REC1
            tmpPayAmt = IIf(IsNumeric(CLM!CLMBordxAmt), CLM!CLMBordxAmt, 0)
            !CasNumber = casno
            !ReceiptNo = txtReceiptNo
            !ReceiptDate = Format(txtReceiptDate, "dd/mm/yyyy")
            !ReceiptAmt = IIf(IsNumeric(txtReceiptAmt), txtReceiptAmt, 0)
            !ChequeNo = ChkStr(txtChequeNo)
            !Remark = ChkStr(TxtRemark)
            !BordxNo = ChkStr(BordxNo)
            !PAYLastUpdateDate = Date
            !PAYLastUpdateUser = Logon
            !RECPAYAmt = tmpPayAmt
            !CheckAmt = tmpPayAmt
            If cboAccPeriodMth <> "" Then
                !RECAccPeriodMth = cboAccPeriodMth
            End If
            If cboAccPeriodYR <> "" Then
                !RECAccPeriodYR = cboAccPeriodYR
            End If
            REC1.update
        End With
        With CLM
            !PVRecPayorStatus = status
        End With
        CLM.update
        REC1.Close
        Set REC1 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateBordxStatus(BordxNo As String, status As String)
Const METHOD_NAME = "UpdateBordxStatus"
On Error GoTo debugs
Dim strsqlBordx As String
    strsqlBordx = "Update Bordx set BordxPAYRecDate='" & Format(txtReceiptDate, "yyyy/mm/dd") & _
                  "', BordxPAYFRStatus='" & status & "' where BordxrefNo = '" & Trim(BordxNo) & "'"
    MediConnCISWS.Execute strsqlBordx
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'*************************** End UpdateRecord

'*************************** Start Option Click
Private Sub OptBordx_Click()
Const METHOD_NAME = "OptBordx_Click"
On Error GoTo debugs

     If OptBordx.Value = True Then
        ChkScan.Value = 1
        lstBordx.Visible = True
        CboBordxStatus.Visible = True
        CboWSStatus.Visible = False
        CboBordxStatus.ListIndex = 0
        CboWSStatus.ListIndex = 0
        txtReceiptNo = ""
        txtReceiptDate = "__/__/____"
        txtChequeNo = ""
        txtReceiptAmt = ""
        TxtRemark = ""
        lstWorksheet.Visible = False
        TxtWSNoTo.Visible = True
        Label21.Visible = True
        txtWSNoFr = ""
        TxtWSNoTo = ""
        ChkScan.Enabled = True
        Label14.Visible = False
        TxtBordx2.Visible = False
        TxtBordx1.SetFocus
        ChkScan.Left = 3360
        ChkScan.Top = 800
     End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub OptWks_Click()
    If OptWks.Value = True Then
        ChkScan.Value = 1
       lstWorksheet.Visible = True
       txtReceiptNo = ""
       txtReceiptDate = "__/__/____"
       txtReceiptAmt = ""
       txtChequeNo = ""
       TxtRemark = ""
       CboBordxStatus.Visible = False
       CboWSStatus.Visible = True
       CboWSStatus.ListIndex = 0
       CboBordxStatus.ListIndex = 0
       lstBordx.Visible = False
       TxtBordx2.Visible = True
       'txtWSNoTo.Enabled = True
       'txtWSNoFr.Enabled = True
       'TxtBordx1.Enabled = False
       'TxtBordx2.Enabled = False
       Label21.Visible = False
       Label14.Visible = True
       TxtWSNoTo.Visible = False
       txtWSNoFr.SetFocus
       ChkScan.Enabled = True
       ChkScan.Value = 1
       ChkScan.Left = 3360
       ChkScan.Top = 530
    End If
End Sub
'*************************** End Option Click

'************** Start ComboListing ****************************
Private Sub ComboListing()
    
    With CboBordxStatus
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
    
    With CboWSStatus
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
    End With
    
    With cboInsured
        .AddItem ""
        .AddItem "YES"
        .AddItem "NO"
    End With
    
End Sub

Private Sub PayorComboListing()
Const METHOD_NAME = "PayorComboListing"
On Error GoTo debugs
    Dim Company As New ADODB.Recordset
    Company.Open "Select GRPCoName from GRPCompany ", MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do Until Company.EOF
            CboPayor.AddItem Trim(Company!GrpCoName)
            Company.MoveNext
        Loop
        Company.Close
        Set Company = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub CIS_PayorComboListing()
Const METHOD_NAME = "CIS_PayorComboListing"
On Error GoTo debugs
    Dim Company As New ADODB.Recordset
    Company.Open "Select PAYCode,PAYCompanyName from CISDB..PAY ", MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do Until Company.EOF
            cisdbcbopayor.AddItem Trim(Company!PAYCode) & " - " & Trim(Company!PAYCompanyName)
            Company.MoveNext
        Loop
        Company.Close
        Set Company = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'************** End ComboListing ****************************

Private Sub txtReceiptAmt_LostFocus()
    If txtReceiptAmt <> lblSelectedAmt Then
        MsgBox "Receipt Amt & Checked Amt Are Different! " & vbNewLine & _
               "Please Fill in Remark.", vbOKOnly + vbCritical
    End If
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
        cboAccPeriodMth.Text = Format(txtReceiptDate, "m")
        cboAccPeriodYR.Text = Format(txtReceiptDate, "yyyy")
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
        End If
    End If
End Sub
Private Sub lstBordx_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(14)
    End Select
End Sub

Private Sub lstWorksheet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstWorksheet, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(16)
    End Select
End Sub

Private Function ClearForm(frmCall)
On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    End If
Next ctl
    txtACPeriod.mask = ""
    txtACPeriod.Text = ""
    txtACPeriod.mask = "##/####"
End Function

Private Function CaseNoCheck()
    If Trim(ChkStr(txtWSNoFr)) = Trim(lstWorksheet.ListItems(listloop1).SubItems(3)) Then
        Check1 = True
    End If
End Function

Private Function BordxNoCheck()
    If Trim(ChkStr(TxtBordx1)) = Trim(lstBordx.ListItems(listloop1).SubItems(2)) Then
        Check2 = True
    End If
End Function

Private Sub TxtReceiptNo_LostFocus()
    txtReceiptNo = ChkStr(txtReceiptNo)
End Sub

Private Sub TxtRemark_lostfocus()
    TxtRemark = ChkStr(TxtRemark)
End Sub

Private Sub txtWSNoFr_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub txtWSNoFr_LostFocus()
    txtWSNoFr = ChkStr(txtWSNoFr)
    CmdSearch.Default = False
End Sub
