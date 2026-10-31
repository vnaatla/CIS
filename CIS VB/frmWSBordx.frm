VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWSBordx 
   Caption         =   "Insert Worksheet for Bordx"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CboBordxtype 
      Height          =   315
      ItemData        =   "frmWSBordx.frx":0000
      Left            =   1440
      List            =   "frmWSBordx.frx":0010
      TabIndex        =   57
      Top             =   2210
      Width           =   1455
   End
   Begin VB.TextBox TxtTempPayor 
      Height          =   315
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   33
      Top             =   8280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox lblBordxNo 
         Height          =   285
         Left            =   5760
         TabIndex        =   61
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3360
         TabIndex        =   60
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1200
         TabIndex        =   59
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   58
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print &File"
         Height          =   375
         Left            =   9000
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdBordx 
         Caption         =   "&Bordx"
         Height          =   375
         Left            =   9720
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   8520
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10320
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10920
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   6960
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7800
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Last Generated Bordx No"
         Height          =   375
         Left            =   4560
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Select Worksheets for Bordx"
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
         TabIndex        =   20
         Top             =   80
         Width           =   3285
      End
   End
   Begin MSComctlLib.ListView lstWorksheetTrackList 
      Height          =   4095
      Left            =   120
      TabIndex        =   32
      Top             =   3120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7223
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
      NumItems        =   40
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Pay To"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Payor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Wks No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bordx Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Non GP Fee"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Date Visit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Patient"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Clinic Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Clinic Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Grp Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Plan Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Treatment Reason"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Invoice Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "OP Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "FFS Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "FFS Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "MCO Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Bill Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Subsidiary"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "InvoiceEnteredAmt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Pay2CLN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "CLN Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "CLMAppStatus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Treatment Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Treatment Reason"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Treatment Reason 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Treatment Reason 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Treatment Reason 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Treatment Reason 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Oth Diagnosis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Doctor Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Co-Pay"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "CoverID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "Claim subsidiary"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "Claim cost center"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "Claim Dept Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame frmWSBordx 
      Height          =   2715
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   11655
      Begin VB.OptionButton optASO 
         Caption         =   "Spl.Bordx"
         Height          =   195
         Index           =   2
         Left            =   4080
         TabIndex        =   68
         Top             =   1920
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.ComboBox cboEntityType 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":0054
         Left            =   1320
         List            =   "frmWSBordx.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   2130
         Width           =   3015
      End
      Begin VB.OptionButton optASO 
         Caption         =   "BTBG"
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   67
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optASO 
         Caption         =   "ASO"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   66
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CboUOB 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":007F
         Left            =   6720
         List            =   "frmWSBordx.frx":0089
         TabIndex        =   64
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox ChkClinic 
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
         Left            =   11040
         TabIndex        =   53
         Top             =   1820
         Width           =   255
      End
      Begin VB.ComboBox CboCLN 
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
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   1820
         Width           =   4180
      End
      Begin VB.ComboBox CboReportType 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":00A3
         Left            =   1320
         List            =   "frmWSBordx.frx":00AD
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   4095
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   440
         Width           =   4095
      End
      Begin VB.CheckBox ChkDOV 
         Caption         =   "Empty Date"
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
         Left            =   10200
         TabIndex        =   11
         Top             =   150
         Width           =   1215
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   2
         Top             =   720
         Width           =   1320
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   315
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   3
         Top             =   720
         Width           =   1320
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   780
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtDateVisit1 
         Height          =   300
         Left            =   6720
         TabIndex        =   6
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateVisit2 
         Height          =   300
         Left            =   8640
         TabIndex        =   7
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtOPBatchNoTo 
         Height          =   315
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1000
         Width           =   1320
      End
      Begin VB.TextBox txtOPBatchNoFr 
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1000
         Width           =   1320
      End
      Begin VB.ComboBox cboCashType 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":00D2
         Left            =   1320
         List            =   "frmWSBordx.frx":00E2
         TabIndex        =   41
         Top             =   1290
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtRegDateFr 
         Height          =   300
         Left            =   6720
         TabIndex        =   45
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRegDateTo 
         Height          =   300
         Left            =   8640
         TabIndex        =   46
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboClinicType 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":0126
         Left            =   6720
         List            =   "frmWSBordx.frx":0136
         TabIndex        =   8
         Top             =   660
         Width           =   4695
      End
      Begin VB.ComboBox cboTreatmentCode 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":017A
         Left            =   6720
         List            =   "frmWSBordx.frx":018A
         TabIndex        =   43
         Top             =   945
         Width           =   4695
      End
      Begin VB.ComboBox cboSubsidiary 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":01CE
         Left            =   6720
         List            =   "frmWSBordx.frx":01D0
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   1230
         Width           =   4695
      End
      Begin VB.ComboBox CboDepartment 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":01D2
         Left            =   6720
         List            =   "frmWSBordx.frx":01D4
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1515
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ComboBox cboOMNIStatus 
         Height          =   315
         ItemData        =   "frmWSBordx.frx":01D6
         Left            =   1320
         List            =   "frmWSBordx.frx":01E6
         TabIndex        =   49
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label lblOPFlex 
         Caption         =   "Outpatient / flexCare"
         Height          =   375
         Left            =   5760
         TabIndex        =   65
         Top             =   2160
         Width           =   975
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
         TabIndex        =   62
         ToolTipText     =   "Date Pass On"
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Bordx Type"
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
         Index           =   3
         Left            =   240
         TabIndex        =   56
         ToolTipText     =   "Date Pass On"
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Clinic"
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
         Left            =   5760
         TabIndex        =   51
         Top             =   1820
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "OMNI"
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
         Index           =   3
         Left            =   240
         TabIndex        =   50
         Top             =   1650
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Index           =   1
         Left            =   8280
         TabIndex        =   48
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "DOR"
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
         Left            =   5760
         TabIndex        =   47
         ToolTipText     =   "Date of Admission"
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Treatment"
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
         Left            =   5760
         TabIndex        =   44
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Top             =   1380
         Width           =   855
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
         Left            =   5760
         TabIndex        =   40
         Top             =   1290
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "To"
         Height          =   255
         Index           =   2
         Left            =   2670
         TabIndex        =   38
         ToolTipText     =   "Date Pass On"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "OP Batch No"
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
         Left            =   240
         TabIndex        =   37
         ToolTipText     =   "Date Pass On"
         Top             =   1075
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Clinic Type"
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
         Index           =   0
         Left            =   5760
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Department"
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
         Left            =   5760
         TabIndex        =   35
         Top             =   1590
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Report Type"
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
         TabIndex        =   34
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Date Visit"
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
         Index           =   0
         Left            =   5760
         TabIndex        =   30
         ToolTipText     =   "Date of Admission"
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Wks. No"
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
         Index           =   0
         Left            =   240
         TabIndex        =   29
         ToolTipText     =   "Date Pass On"
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Index           =   0
         Left            =   8280
         TabIndex        =   28
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   2670
         TabIndex        =   27
         Top             =   795
         Width           =   375
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Clinic"
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
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Clinic"
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
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmWSBordx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intExit As Integer
Dim CaseNoArray() As String
Dim CLMBordxNo As String
Dim casno As String
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim ScanCode As String
Dim PLNCode As String
Dim bIsBTBG As Boolean
Dim strtmpFlexSelection As String
Dim strUBAppend As String
Private m_blnDirection(1 To 40) As Boolean


Private Sub CboBordxtype_Change()
   'this method is not working
End Sub

Private Sub CboBordxtype_Click()
 If CboBordxtype.ListIndex = 1 Then
        optASO(0).Visible = True
        optASO(1).Visible = True
        optASO(2).Visible = True
    Else
        optASO(0).Visible = False
        optASO(1).Visible = False
        optASO(2).Visible = False
    End If
End Sub


Private Sub cboCLN_Change()
If InStr(CboCLN.Text, "null") Then
       CboCLN.Enabled = False
    End If
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
    
      Dim GrpCompanyName As String

         GrpCompanyName = Replace(CboPayor, "'", "''")
        If Len(GrpCompanyName) = 0 Then
            Exit Sub
        Else
            GrpCompanyName = Split(GrpCompanyName, "|")(0)
        End If
   ' sql = " Select GRPCoPAYCode, GRPCoName From GRPCompany Where GRPCoName = '" & Trim(GRPCOMName) & "'"
   '
   ' rst2.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    
   'If rst2.EOF = False Then '.RecordCount Then
    '   TxtTempPayor = IIf(Len(rst2!GRPCoPAYCode), rst2!GRPCoPAYCode, "")
        TxtTempPayor = "DC"
   'End If
   ' rst2.Close
   ' Set rst2 = Nothing
    
    If CboReportType.ListIndex = 1 Or CboReportType.ListIndex = 0 Then
        '=Run Sp=================================================
        'GRPCOMName
        sql2 = "Exec SearchDepartment '" & Trim(GrpCompanyName) & "'"
        rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        '===================================================
        CboDepartment.Clear
        Do Until rst3.EOF
            With rst3
                CboDepartment.AddItem IIf(Len(Trim(!MBMDepartment)), Trim(!MBMDepartment), "")
                rst3.MoveNext
            End With
        Loop
        rst3.Close
        Set rst3 = Nothing
    End If
    
        '=Run Sp=================================================
      
        
        
        sql2 = ""
        sql2 = "Exec SearchSubsidiary_bygrp '" & Trim(GrpCompanyName) & "'"
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

Private Sub CboReportType_Change()

    If CboReportType.ListIndex = 1 Then
        CboDepartment.Visible = True
        Label6.Visible = True
    Else
        CboDepartment.Visible = False
        Label6.Visible = False
    End If
        
    If InStr(CboReportType.Text, "null") Then
        CboReportType.Enabled = False
    End If

End Sub

Private Sub CboReportType_Click()

    If CboReportType.ListIndex = 1 Then
        CboDepartment.Visible = True
        Label6.Visible = True
    Else
        CboDepartment.Visible = True
        Label6.Visible = True
    End If

End Sub


'# Added by Ravi
Private Sub CboUOB_Click()
    'Selection: Cashless  - Out Patient
    If cboCashType.ListIndex = 0 And CboUOB.ListIndex = 0 Then
            strUBAppend = ""
            lstWorksheetTrackList.ListItems.Clear
            Call SearchFlexOPRecord
            Call WorksheetopTrackingListingUB(strUBAppend)
    'Selection: Cashless  - Flexcare
    ElseIf cboCashType.ListIndex = 0 And CboUOB.ListIndex = 1 Then
            strUBAppend = ""
            lstWorksheetTrackList.ListItems.Clear
            Call SearchFlexOPRecord
            Call WorksheetflexTrackingListingUB(strUBAppend)
    'Selection: Reim  - Out Patient
    ElseIf cboCashType.ListIndex = 1 And CboUOB.ListIndex = 0 Then
            strUBAppend = ""
            lstWorksheetTrackList.ListItems.Clear
            Call SearchFlexOPRecord
            Call WorksheetopTrackingListingUB(strUBAppend)
    'Selection: Reim  - FlexCare
    ElseIf cboCashType.ListIndex = 1 And CboUOB.ListIndex = 1 Then
            strUBAppend = ""
            lstWorksheetTrackList.ListItems.Clear
            Call SearchFlexOPRecord
            Call WorksheetflexTrackingListingUB(strUBAppend)
    End If
    If Trim(CboUOB) = "" Then
        MsgBox "Please Select Search Criteria!", vbCritical + vbOKOnly, DialogBoxTitle
        CboUOB.SetFocus
    End If
End Sub


'****************** Start WorksheetflexTrackingListingUB *****************************
  'change done by ravi 21102019
Private Sub WorksheetflexTrackingListingUB(Optional strUBAppend As String)
Const METHOD_NAME = "WorksheetflexTrackingListingUB"
'On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim strAppend2 As String
 
    total = 0
    Current = 0
    lstWorksheetTrackList.ListItems.Clear
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 500
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strUBAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, 23)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("BordxType", adTinyInt, adParamInput, 1, CboBordxtype.ListIndex)
    cmd.Parameters.Append Prm3
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdBordx.Enabled = True
        lblCurrent = 0
        lblTotal = 0
        LblTotalAmt = 0
    Else
    Do
        Do Until rst2.EOF
        total = rst2.RecordCount
        lblTotal = total
        
            With rst2
                Set Record = lstWorksheetTrackList.ListItems.Add(, , "")
        
                Record.SubItems(1) = IIf(Len(Trim(!INVPayTo)), Trim(!INVPayTo), "")
                Record.SubItems(2) = IIf(Len(Trim(!PAYCode)), Trim(!PAYCode), "")
                Record.SubItems(3) = IIf(Len(Trim(!CasNumber)), Trim(!CasNumber), "")
                Record.SubItems(4) = IIf(IsNumeric(!CLMBordxAmt), Format(!CLMBordxAmt, "#,##0.00"), 0)
                Record.SubItems(5) = IIf(IsNumeric(!ClmNonGPFee), Format(!ClmNonGPFee, "#,##0.00"), 0)
                If Len(!DateVisit) Then
                   Record.SubItems(6) = Format(!DateVisit, "dd/mm/yyyy")
                End If
                Record.SubItems(7) = IIf(Len(Trim(!MBMNumber)), Trim(!MBMNumber), "")
                If !coverid = "00" Then
                    Record.SubItems(8) = IIf(Len(Trim(!MBMName)), Trim(!MBMName), "")
                Else
                    Record.SubItems(8) = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                End If
                Record.SubItems(9) = IIf(Len(Trim(!ClinicCode)), Trim(!ClinicCode), "")
                Record.SubItems(10) = IIf(Len(Trim(!CLNName)), Trim(!CLNName), "")
                Record.SubItems(11) = IIf(Len(Trim(!MBMGroupCompany)), Trim(!MBMGroupCompany), "")
                Record.SubItems(12) = IIf(Len(Trim(!PLNCode)), Trim(!PLNCode), "")
                Record.SubItems(13) = IIf(Len(Trim(!ReasonDesc1)), Trim(!ReasonDesc1), "")
                If !MBMHISPolNo = "03215666" Then
                    Record.SubItems(14) = "AX - AXA AFFIN (NON-EXECUTIVE)"
                ElseIf !MBMHISPolNo = "03215661" Then
                    Record.SubItems(14) = "AX - AXA AFFIN (EXECUTIVE AND ABOVE)"
                Else
                    Record.SubItems(14) = IIf(Len(Trim(!MBMPosition)), Trim(!MBMPosition), "")
                End If
                Record.SubItems(15) = IIf(Len(Trim(!InvoiceNo)), Trim(!InvoiceNo), "")
                Record.SubItems(16) = IIf(Len(Trim(!OPBatchNo)), Trim(!OPBatchNo), "")
                Record.SubItems(17) = IIf(Len(Trim(!CLNFFSStatus)), Trim(!CLNFFSStatus), "")
                If Len(Trim(!CLMMCOProcFee)) Then
                    If InStr(1, !CLMMCOProcFee, "-") Then
                     Record.SubItems(18) = Format(!CLMMCOProcFee, "#,##0.00") * (-1)
                    Else
                     Record.SubItems(18) = Format(!CLMMCOProcFee, "#,##0.00")
                    End If
                Else
                     Record.SubItems(18) = 0
                End If
                Record.SubItems(19) = IIf(Len(Trim(!GRPMCOStatus)), !GRPMCOStatus, "")
                Record.SubItems(20) = IIf(Len(Trim(!CLMTotalActual)), Format(!CLMTotalActual, "#,##0.00"), "")
                Record.SubItems(21) = IIf(Len(Trim(!Subsidiary)), Trim(!Subsidiary), "")
                Record.SubItems(22) = IIf(Len(Trim(!INVBillAmt)), Format(!INVBillAmt, "#,##0.00"), "")
                Record.SubItems(23) = IIf(Len(Trim(!CLMPayableByMedix2c)), Format(!CLMPayableByMedix2c, "#,##0.00"), "")
                Record.SubItems(24) = IIf(Len(Trim(!CLNType)), !CLNType, "")
                Record.SubItems(25) = IIf(Len(Trim(!CASClaimsApp)), !CASClaimsApp, "")
                Record.SubItems(27) = IIf(Len(Trim(!ReasonDesc1)), !ReasonDesc1, "")
                Record.SubItems(28) = IIf(Len(Trim(!ReasonDesc2)), !ReasonDesc2, "")
                Record.SubItems(29) = IIf(Len(Trim(!ReasonDesc3)), !ReasonDesc3, "")
                Record.SubItems(30) = IIf(Len(Trim(!ReasonDesc4)), !ReasonDesc4, "")
                Record.SubItems(31) = IIf(Len(Trim(!ReasonDesc5)), !ReasonDesc5, "")
                Record.SubItems(32) = IIf(Len(Trim(!CASDiagnosis5)), !CASDiagnosis5, "")
                Record.SubItems(33) = IIf(Len(Trim(!MBMPolicyNo)), !MBMPolicyNo, "")
                Record.SubItems(34) = IIf(Len(Trim(!DoctorName)), !DoctorName, "")
                Record.SubItems(35) = IIf(Len(Trim(!CLMCoPayColl)), Format(!CLMCoPayColl, "#,##0.00"), "")
                Record.SubItems(36) = IIf(Len(Trim(!coverid)), Trim(!coverid), "")
                Record.SubItems(37) = IIf(Len(Trim(!clmsubsidiary)), Trim(!clmsubsidiary), "")
                Record.SubItems(38) = IIf(Len(Trim(!clmcostcenter)), Trim(!clmcostcenter), "")
                Record.SubItems(39) = IIf(Len(Trim(!CLMDeptName)), Trim(!CLMDeptName), "")
                
                Current = Current + 1
                lblCurrent = Current
                
                If Len(!CLMBordxAmt) Then
                    TotalAmt = TotalAmt + !CLMBordxAmt
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
    CmdExit.Enabled = True
    CmdBordx.Enabled = True
    
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'# Added by Ravi


'change done by ravi 21102019
Private Function SearchFlexOPRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim StrSql As String
Dim strAppend As String
Dim sql As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String
Dim Claim As New ADODB.Recordset
Dim GrpCompanyName As String
    strAppend = ""
  
    
    GrpCompanyName = ""
   lstWorksheetTrackList.ListItems.Clear
    If Len(Trim(CboPayor)) Then
         GrpCompanyName = Replace(CboPayor, "'", "''")
         GrpCompanyName = Split(GrpCompanyName, "|")(1)
              
         'strAppend = strAppend + " AND HISDB.dbo.MBMOthers.MBMGroupCompany = '" & Trim(GrpCompanyName) & "'"
         strAppend = strAppend + " AND CISWorksheet.dbo.Claim.CLMGrpComCode = '" & Trim(GrpCompanyName) & "'"
    End If
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtCaseNo1)) Then
            strAppend = strAppend + " AND CISDB.dbo.CISCas.CASNumber = '" & Trim(TxtCaseNo1) & "'"
        End If
    Else
        If Len(Trim(TxtCaseNo1)) Then
            strAppend = strAppend + " AND CISDB.dbo.CISCas.CASNumber >= '" & Trim(TxtCaseNo1) & "'"
        End If
        
        If Len(Trim(TxtCaseNo2)) Then
            strAppend = strAppend + " AND CISDB.dbo.CISCas.CASNumber <= '" & Trim(TxtCaseNo2) & "'"
        End If
    End If
    
    If Len(Trim(txtDateVisit1)) > 0 And IsDate(txtDateVisit1) = True Then
        strAppend = strAppend + " And CISCas.DateVisit >= '" & Format(txtDateVisit1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateVisit2)) > 0 And IsDate(txtDateVisit2) = True Then
        strAppend = strAppend + " And CISCas.DateVisit <= '" & Format(txtDateVisit2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkDOV.Value = 1 Then
        sql = ""
        sql = "AND (CISCas.DateVisit Is null OR CISCas.DateVisit = '')" & ""
        strAppend = strAppend + sql
    End If
    If Len(Trim(CboDepartment)) Then
        strAppend = strAppend + " AND MBMOthers.MBMDepartment = '" & Trim(CboDepartment) & "'"
    End If
    If Mid(CboClinicType, 1, 1) = 1 Then
        strAppend = strAppend & " AND CLN.CLNType = 'S'"
    ElseIf Mid(CboClinicType, 1, 1) = 2 Then
        strAppend = strAppend & " AND CLN.CLNType = 'G'"
    ElseIf Mid(CboClinicType, 1, 1) = 3 Then
        strAppend = strAppend & " AND  (CLN.CLNType <> 'G') AND (CLN.CLNType <> 'S')"
    End If
    If Len(Trim(txtOPBatchNoFr)) Then
        strAppend = strAppend + " AND CISCas.OPBatchNo >= '" & Trim(txtOPBatchNoFr) & "'"
    End If
    
    If Len(Trim(txtOPBatchNoTo)) Then
        strAppend = strAppend + " AND CISCas.OPBatchNo <= '" & Trim(txtOPBatchNoTo) & "'"
    End If
    
    If Len(Trim(cboSubsidiary)) Then
        strAppend = strAppend + " AND MBMOthers.Subsidiary = '" & Trim(cboSubsidiary) & "'"
    End If
    
    If Len(Trim(cboCashType)) Then
        If Mid(cboCashType, 1, 1) = "C" Then
            strAppend = strAppend + " AND CISWorksheet.dbo.ClaimInvoice.INVPayto = 'C'"
        Else
            strAppend = strAppend + " AND CISWorksheet.dbo.ClaimInvoice.INVPayto = 'M'"
        End If
    End If
    
    If Len(Trim(cboTreatmentCode)) Then
        strAppend = strAppend + " AND CASTreatReason1 = '" & Mid(cboTreatmentCode, 1, 3) & "'"
    End If
    
    If Len(Trim(cboOMNIStatus)) Then
        strAppend = strAppend + " AND GRPCompany.GRPOMNICLM = 'Y'"
    End If
    
    If Len(Trim(txtRegDateFr)) > 0 And IsDate(txtRegDateFr) = True Then
        strAppend = strAppend + " And CISCas.RegisterDate >= '" & Format(txtRegDateFr, "yyyy/MM/DD") & "'"
    End If
    
    If Len(Trim(txtRegDateTo)) > 0 And IsDate(txtRegDateTo) = True Then
        strAppend = strAppend + " And CISCas.RegisterDate <= '" & Format(txtRegDateTo, "yyyy/MM/DD") & "'"
    End If
    If Len(Trim(CboCLN)) Then
        strAppend = strAppend + " AND CISCas.ClinicCode = '" & Trim(Left(CboCLN, IIf(InStr(1, CboCLN, "-") - 1 < 0, Len(CboCLN), InStr(1, CboCLN, "-") - 1))) & "'"
    End If
    
                         
    If cboEntityType.ListIndex = 0 Then
        strAppend = strAppend + " AND Substring(CISCAS.CasNumber, 1, 2) <>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        strAppend = strAppend + " AND Substring(CISCAS.CasNumber, 1, 2) ='SJ' "
    End If
    
  '#change done by ravi 30102019
     'Selection: Cashless  - op
    If cboCashType.ListIndex = 0 And CboUOB.ListIndex = 0 Then
            strUBAppend = ""
            'strUBAppend = strAppend + "AND (tblFlexibleBenefitPlan_OP.FBEmployee is null or (tblFlexibleBenefitPlan_OP.FBFamily =1 and CISCas.CoverID<>'00')or tblFlexibleBenefitPlan_OP.FBisOPFlexCombine = 1)"
            strUBAppend = strAppend + " AND ((tblFlexibleBenefitPlanLimit_OP.CISPlnCode is NOT null and tblFlexibleBenefitPlan_OP.FBCode is null) OR (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is null) or (tblFlexibleBenefitPlan_OP.FBFamily=1 and ciscas.CoverID<>'00') or tblFlexibleBenefitPlan_OP.FBisOPFlexCombine = 1 )"
    'Selection: Cashless  - Flexcare
    ElseIf cboCashType.ListIndex = 0 And CboUOB.ListIndex = 1 Then
            strUBAppend = ""
            strUBAppend = strAppend + " AND (tblFlexibleBenefitPlan_OP.FBEmployee is not null and CISCas.CoverID='00') AND (tblFlexibleBenefitPlan_OP.FBisOPFlexCombine <> 1) "
    'Selection: Reim  - Out Patient
    ElseIf cboCashType.ListIndex = 1 And CboUOB.ListIndex = 0 Then
            strUBAppend = ""
            'strUBAppend = strAppend + " AND (tblFlexibleBenefitPlan_OP.FBEmployee is null or (tblFlexibleBenefitPlan_OP.FBFamily =1 and CISCas.CoverID<>'00') or (tblFlexibleBenefitPlan_OP.FBCode Not in ('305,303') AND MBMCLNOthers.PLNType IN('U131','U139','U132','U150','U151','U152','U153','U154','U155')))"
            strUBAppend = strAppend + " AND ((tblFlexibleBenefitPlanLimit_OP.CISPlnCode is NOT null and tblFlexibleBenefitPlan_OP.FBCode is null) OR (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is null) or (tblFlexibleBenefitPlan_OP.FBFamily=1 and ciscas.CoverID<>'00'))"
    'Selection: Reim  - FlexCare
    ElseIf cboCashType.ListIndex = 1 And CboUOB.ListIndex = 1 Then
            strUBAppend = ""
            'strUBAppend = strAppend + " AND (tblFlexibleBenefitPlan_OP.FBEmployee is not null and CISCas.CoverID='00') "
            strUBAppend = strAppend + " AND (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null and tblFlexibleBenefitPlan_OP.FBEmployee=1 and ciscas.CoverID='00')"
    End If
    
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Function
  'change done by ravi 21102019
Private Sub WorksheetopTrackingListingUB(Optional strUBAppend As String)
Const METHOD_NAME = "WorksheetopTrackingListingUB"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim strAppend2 As String
 
    total = 0
    Current = 0
    lstWorksheetTrackList.ListItems.Clear
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 500
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strUBAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, 24)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("BordxType", adTinyInt, adParamInput, 1, CboBordxtype.ListIndex)
    cmd.Parameters.Append Prm3
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdBordx.Enabled = True
        lblCurrent = 0
        lblTotal = 0
        LblTotalAmt = 0
    Else
    Do
        Do Until rst2.EOF
        total = rst2.RecordCount
        lblTotal = total
        
            With rst2
                Set Record = lstWorksheetTrackList.ListItems.Add(, , "")
        
                Record.SubItems(1) = IIf(Len(Trim(!INVPayTo)), Trim(!INVPayTo), "")
                Record.SubItems(2) = IIf(Len(Trim(!PAYCode)), Trim(!PAYCode), "")
                Record.SubItems(3) = IIf(Len(Trim(!CasNumber)), Trim(!CasNumber), "")
                Record.SubItems(4) = IIf(IsNumeric(!CLMBordxAmt), Format(!CLMBordxAmt, "#,##0.00"), 0)
                Record.SubItems(5) = IIf(IsNumeric(!ClmNonGPFee), Format(!ClmNonGPFee, "#,##0.00"), 0)
                If Len(!DateVisit) Then
                   Record.SubItems(6) = Format(!DateVisit, "dd/mm/yyyy")
                End If
                Record.SubItems(7) = IIf(Len(Trim(!MBMNumber)), Trim(!MBMNumber), "")
                If !coverid = "00" Then
                    Record.SubItems(8) = IIf(Len(Trim(!MBMName)), Trim(!MBMName), "")
                Else
                    Record.SubItems(8) = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                End If
                Record.SubItems(9) = IIf(Len(Trim(!ClinicCode)), Trim(!ClinicCode), "")
                Record.SubItems(10) = IIf(Len(Trim(!CLNName)), Trim(!CLNName), "")
                Record.SubItems(11) = IIf(Len(Trim(!MBMGroupCompany)), Trim(!MBMGroupCompany), "")
                Record.SubItems(12) = IIf(Len(Trim(!PLNCode)), Trim(!PLNCode), "")
                Record.SubItems(13) = IIf(Len(Trim(!ReasonDesc1)), Trim(!ReasonDesc1), "")
                If !MBMHISPolNo = "03215666" Then
                    Record.SubItems(14) = "AX - AXA AFFIN (NON-EXECUTIVE)"
                ElseIf !MBMHISPolNo = "03215661" Then
                    Record.SubItems(14) = "AX - AXA AFFIN (EXECUTIVE AND ABOVE)"
                Else
                    Record.SubItems(14) = IIf(Len(Trim(!MBMPosition)), Trim(!MBMPosition), "")
                End If
                Record.SubItems(15) = IIf(Len(Trim(!InvoiceNo)), Trim(!InvoiceNo), "")
                Record.SubItems(16) = IIf(Len(Trim(!OPBatchNo)), Trim(!OPBatchNo), "")
                Record.SubItems(17) = IIf(Len(Trim(!CLNFFSStatus)), Trim(!CLNFFSStatus), "")
                If Len(Trim(!CLMMCOProcFee)) Then
                    If InStr(1, !CLMMCOProcFee, "-") Then
                     Record.SubItems(18) = Format(!CLMMCOProcFee, "#,##0.00") * (-1)
                    Else
                     Record.SubItems(18) = Format(!CLMMCOProcFee, "#,##0.00")
                    End If
                Else
                     Record.SubItems(18) = 0
                End If
                Record.SubItems(19) = IIf(Len(Trim(!GRPMCOStatus)), !GRPMCOStatus, "")
                Record.SubItems(20) = IIf(Len(Trim(!CLMTotalActual)), Format(!CLMTotalActual, "#,##0.00"), "")
                Record.SubItems(21) = IIf(Len(Trim(!Subsidiary)), Trim(!Subsidiary), "")
                Record.SubItems(22) = IIf(Len(Trim(!INVBillAmt)), Format(!INVBillAmt, "#,##0.00"), "")
                Record.SubItems(23) = IIf(Len(Trim(!CLMPayableByMedix2c)), Format(!CLMPayableByMedix2c, "#,##0.00"), "")
                Record.SubItems(24) = IIf(Len(Trim(!CLNType)), !CLNType, "")
                Record.SubItems(25) = IIf(Len(Trim(!CASClaimsApp)), !CASClaimsApp, "")
                Record.SubItems(27) = IIf(Len(Trim(!ReasonDesc1)), !ReasonDesc1, "")
                Record.SubItems(28) = IIf(Len(Trim(!ReasonDesc2)), !ReasonDesc2, "")
                Record.SubItems(29) = IIf(Len(Trim(!ReasonDesc3)), !ReasonDesc3, "")
                Record.SubItems(30) = IIf(Len(Trim(!ReasonDesc4)), !ReasonDesc4, "")
                Record.SubItems(31) = IIf(Len(Trim(!ReasonDesc5)), !ReasonDesc5, "")
                Record.SubItems(32) = IIf(Len(Trim(!CASDiagnosis5)), !CASDiagnosis5, "")
                Record.SubItems(33) = IIf(Len(Trim(!MBMPolicyNo)), !MBMPolicyNo, "")
                Record.SubItems(34) = IIf(Len(Trim(!DoctorName)), !DoctorName, "")
                Record.SubItems(35) = IIf(Len(Trim(!CLMCoPayColl)), Format(!CLMCoPayColl, "#,##0.00"), "")
                Record.SubItems(36) = IIf(Len(Trim(!coverid)), Trim(!coverid), "")
                Record.SubItems(37) = IIf(Len(Trim(!clmsubsidiary)), Trim(!clmsubsidiary), "")
                Record.SubItems(38) = IIf(Len(Trim(!clmcostcenter)), Trim(!clmcostcenter), "")
                Record.SubItems(39) = IIf(Len(Trim(!CLMDeptName)), Trim(!CLMDeptName), "")
                
                Current = Current + 1
                lblCurrent = Current
                
                If Len(!CLMBordxAmt) Then
                    TotalAmt = TotalAmt + !CLMBordxAmt
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
    CmdExit.Enabled = True
    CmdBordx.Enabled = True
    
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'# Added by Ravi
Private Sub ChkClinic_Click()
Dim CLNRec As New ADODB.Recordset

    CboCLN.Clear
    If ChkClinic.Value = 1 Then
        CLNRec.Open "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do Until CLNRec.EOF
            If Len(CLNRec!CLNName) Then
                CboCLN.AddItem CLNRec!CLNCode & " - " & Trim(CLNRec!CLNName)
            End If
            CLNRec.MoveNext
        Loop
        CLNRec.Close
        Set CLNRec = Nothing
        End If
End Sub

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstWorksheetTrackList.ListItems.Clear
        ChkScan.Left = 3120
        Label2.Visible = False
        TxtCaseNo2.Visible = False
        TxtCaseNo1.SetFocus
    Else
        lstWorksheetTrackList.ListItems.Clear
        ChkScan.Left = 5040
        Label2.Visible = True
        TxtCaseNo2.Visible = True
        TxtCaseNo1.SetFocus
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

Private Sub Form_Load()
Call ActionLogUpdate(12, "frmWSBordx")
    Check = False
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    CboClinicType.ListIndex = 0
    CboReportType.ListIndex = 0
    Call PayorComboListing
    ReDim CaseNoArray(0)
    
    cboEntityType.ListIndex = 0
    CboUOB.Visible = False
    lblOPFlex.Visible = False
      
    
End Sub

'****************** Start Command Button *****************************
Private Sub cmdPrint_Click()
Const METHOD_NAME = "CmdPrint_Click"
On Error GoTo debugs

    Screen.MousePointer = vbHourglass
    
    If lstWorksheetTrackList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstWorksheetTrackList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdBordx_Click()
Const METHOD_NAME = "CmdBordx_Click"
On Error GoTo debugs
'On Error GoTo errSave
Dim PAYCode As String
Dim ii As Integer
Dim Cnt As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim update
Dim PayorIndex As Integer
Dim startTrans
Dim TotalApp As Variant
Dim TotCLM As Integer


    Cnt = 0
    startTrans = False
    bIsBTBG = False
    
    TotalApp = 0
    TotCLM = 0
        
        If lstWorksheetTrackList.ListItems.Count = 0 Then
            MsgBox "Please select a record before proceeding!", vbOKOnly + vbCritical
            TxtCaseNo1.SetFocus
        ElseIf UBound(CaseNoArray) = 0 Then
            MsgBox "Please select at least one record to be updated! ", vbCritical
            TxtCaseNo1.SetFocus
        ElseIf Len(CboBordxtype) = 0 Then
             MsgBox "Please select Bordx Type!", vbOKOnly + vbCritical
             CboBordxtype.SetFocus
       '# Added by Ravi
        ElseIf CboUOB.Visible = True And Len(CboUOB) = 0 Then
            MsgBox "Please select Flexcare or Outpatient Type!", vbOKOnly + vbCritical
            CboBordxtype.SetFocus
       '# Added by Ravi
        Else
            update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
             If update = vbYes Then
                startTrans = True
                If CboBordxtype.ListIndex <> "0" Then
                    bIsBTBG = CheckIsBTBG(PLNCode)
                Else
                    bIsBTBG = True
                End If
                
                MediConnCISWS.beginTrans

                If CboBordxtype.ListIndex = "0" Then
                    PAYCode = "IS"
                ElseIf CboBordxtype.ListIndex = "1" Then
                    PAYCode = "DC"
                ElseIf CboBordxtype.ListIndex = "2" Then
                    PAYCode = "CN"
                ElseIf CboBordxtype.ListIndex = "3" And cboEntityType.ListIndex = 1 Then 'SIHAT JOHOR
                    PAYCode = "SJ"
                ElseIf CboBordxtype.ListIndex = "4" Then
                    PAYCode = "IV"
                Else
                    MsgBox "Please select Bordx Type!", vbOKOnly + vbCritical
                    TxtCaseNo1.SetFocus
                End If

                If CboBordxtype.ListIndex = "4" Then
                    CLMBordxNo = GenerateCLMBordxNoVOID(PAYCode)
                Else
                    CLMBordxNo = GenerateCLMBordxNo(PAYCode)
                End If

                If (PAYCode = "DC" And bIsBTBG = True) Or (PAYCode = "CN" And bIsBTBG = True) Then
                    For ii = 1 To UBound(CaseNoArray)
                     If Trim(CaseNoArray(ii)) <> "" Then
                        casno = Trim(Mid(CaseNoArray(ii), 1, IIf(InStr(1, CaseNoArray(ii), "-") = 0, 13, InStr(1, CaseNoArray(ii), "-") - 1)))
    
                        strsqlCLM = "Select CASNumber, coalesce(CLMBTBGAmt,0) As CLMBTBGAmt From Claim Where CASNumber ='" & Trim(casno) & "'"
                        CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
                        Do While Not CLM.EOF
                            TotalApp = TotalApp + CLM!CLMBTBGAmt
                            TotCLM = TotCLM + 1
                            CLM.MoveNext
                        Loop
                        CLM.Close
                        Set CLM = Nothing
                        Call UpdateBTBGRecord
                     End If
                   Next ii
                ElseIf PAYCode = "IV" Then
                    For ii = 1 To UBound(CaseNoArray)
                     If Trim(CaseNoArray(ii)) <> "" Then
                        casno = Trim(Mid(CaseNoArray(ii), 1, IIf(InStr(1, CaseNoArray(ii), "-") = 0, 13, InStr(1, CaseNoArray(ii), "-") - 1)))
    
                        strsqlCLM = "Select CASNumber, CLMFinalApprove From Claim Where CASNumber ='" & Trim(casno) & "'"
                        CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
                        Do While Not CLM.EOF
                            TotalApp = TotalApp + CLM!CLMFinalApprove
                            TotCLM = TotCLM + 1
                            CLM.MoveNext
                        Loop
                        CLM.Close
                        Set CLM = Nothing
                        Call UpdateRecordvoid
                     End If
                   Next ii
                Else
                    For ii = 1 To UBound(CaseNoArray)
                     If Trim(CaseNoArray(ii)) <> "" Then
                        casno = Trim(Mid(CaseNoArray(ii), 1, IIf(InStr(1, CaseNoArray(ii), "-") = 0, 13, InStr(1, CaseNoArray(ii), "-") - 1)))
    
                        strsqlCLM = "Select CASNumber, CLMFinalApprove From Claim Where CASNumber ='" & Trim(casno) & "'"
                        CLM.Open strsqlCLM, MediConnCISWS, adOpenKeyset, adLockOptimistic
                        Do While Not CLM.EOF
                            TotalApp = TotalApp + CLM!CLMFinalApprove
                            TotCLM = TotCLM + 1
                            CLM.MoveNext
                        Loop
                        CLM.Close
                        Set CLM = Nothing
                        Call UpdateRecord
                     End If
                   Next ii
               End If
               
               Erase CaseNoArray
               ReDim CaseNoArray(0)

               Call SaveCLMBordx(TotalApp, TotCLM)

                If startTrans = True Then
                    MediConnCISWS.commitTrans
                    startTrans = False
                End If
                lstWorksheetTrackList.ListItems.Clear
                LblTotalAmt = 0
                lblTotal = 0
                lblCurrent = 0

                MsgBox "Record Saved, Your Bordx No is : " & CLMBordxNo, vbOKOnly + vbInformation, DialogBoxTitle
                lblBordxNo = CLMBordxNo
                TxtCaseNo1.SetFocus
             End If
        End If
    If startTrans = True Then
        MediConnCISWS.RollbackTrans
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function CheckIsBTBG(strPlnCode As String) As Boolean
Dim i As Integer
Dim MRcmd As New ADODB.Command
Dim MRPrm1 As New ADODB.Parameter
Dim MRRes As New ADODB.Recordset
Dim sRetValue As Boolean

    CheckIsBTBG = False
                
    Set MRcmd.ActiveConnection = MediConnCISDB
    MRcmd.CommandText = "sp_GetPLNBTBGInd"
    MRcmd.CommandType = adCmdStoredProc
    MRcmd.CommandTimeout = 300
    Set MRPrm1 = MRcmd.CreateParameter("PLNCode", adVarChar, adParamInput, 8, strPlnCode)
    MRcmd.Parameters.Append MRPrm1
    MRcmd.Parameters.Append MRcmd.CreateParameter("RetOut", adBoolean, adParamInputOutput, 1, 0)
    MRRes.CursorLocation = adUseClient
    MRRes.CursorType = adOpenForwardOnly
    MRRes.CacheSize = 100
    Set MRRes = MRcmd.Execute
    CheckIsBTBG = MRcmd.Parameters("RetOut")
    Set MRPrm1 = Nothing
    Set MRcmd = Nothing
    Set MRRes = Nothing
End Function

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CboPayor.Clear
    CboClinicType.ListIndex = 0
    lstWorksheetTrackList.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    lblBordxNo = ""
    TxtCaseNo1.SetFocus
    Check = False
    ReDim CaseNoArray(0)
    Call PayorComboListing
    optASO(0).Value = False
    optASO(1).Value = False
    optASO(2).Value = False
    optASO(0).Visible = False
    optASO(1).Visible = False
    optASO(2).Visible = False
End Sub

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    Check = False
    TxtCaseNo1.SetFocus
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Public Sub cmdSearch_Click()
Const METHOD_NAME = "CmdSearch_Click"
On Error GoTo debugs

Dim listrecord As Integer
ScanCode = TxtCaseNo1
If Mid(ScanCode, 1, 2) = "LO" Then
    TxtCaseNo1 = ""
Else
    
    If ChkScan.Value = 0 Then
        lblCurrent = 0
        lblTotal = 0
        LblTotalAmt = 0
    End If
    lblBordxNo = ""
    ReDim CaseNoArray(0)
    If ChkScan.Value = 1 Then
        If CboPayor = "" Then
            MsgBox "Please select a Group Company Name", vbOKOnly + vbCritical, DialogBoxTitle
            CboPayor.SetFocus
        ElseIf TxtCaseNo1 = "" Then
            MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
            TxtCaseNo1.SetFocus
        ElseIf Len(CboBordxtype) = 0 Then
             MsgBox "Please select Bordx Type!", vbOKOnly + vbCritical
            CboBordxtype.SetFocus
        Else
            If CboBordxtype.ListIndex = 1 And (optASO(0).Value = False And optASO(1).Value = False And optASO(2).Value = False) Then
                 MsgBox "Please select Bordx Type whether ASO or BTBG..!", vbOKOnly + vbCritical
                 Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            listrecord = lstWorksheetTrackList.ListItems.Count
            If listrecord <> 0 Then
                For listloop1 = 1 To listrecord
                    Call CaseNoCheck
                Next listloop1
            End If
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    CmdClear.Enabled = False
                    CmdExit.Enabled = False
                    CmdBordx.Enabled = False
                    Call SearchRecord
                End If
                Check = False
            Screen.MousePointer = Default
            TxtCaseNo1 = ""
            TxtCaseNo1.SetFocus
        End If
        TxtCaseNo1.SetFocus
    Else
        If Len(CboBordxtype) = 0 Then
             MsgBox "Please select Bordx Type!", vbOKOnly + vbCritical
            CboBordxtype.SetFocus
        Else
            If CboBordxtype.ListIndex = 1 And (optASO(0).Value = False And optASO(1).Value = False And optASO(2).Value = False) Then
                 MsgBox "Please select Bordx Type whether ASO or BTBG..!", vbOKOnly + vbCritical
                 Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            Call SearchRecord
            Screen.MousePointer = Default
            TxtCaseNo1.SetFocus
        End If
    End If
End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
'****************** End Command Button *****************************
'**********************Start Sorted ListView *******************************
Private Sub lstWorksheetTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
    Case 20
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
    Case 22
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
    Case 23
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
    Case 24
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
    Case 25
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(25)
    Case 26
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(26)
    Case 27
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(27)
    Case 28
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(28)
    Case 29
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(29)
    Case 30
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(30)
    Case 31
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(31)
    Case 32
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(32)
    Case 33
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(33)
    Case 34
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(34)
    Case 35
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(35)
    Case 36
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(36)
    Case 37
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(37)
    Case 38
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(38)
    Case 39
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(39)
    Case 40
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(40)

    End Select
End Sub
'**********************End Sorted ListView *******************************

'****************** Start Search Record *****************************

Private Function SearchRecord()
Const METHOD_NAME = "SearchRecord"
On Error GoTo debugs

Dim StrSql As String
Dim strAppend As String
Dim sql As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String
Dim Claim As New ADODB.Recordset
Dim GrpCompanyName As String
    strAppend = ""
  
    
    GrpCompanyName = ""
   lstWorksheetTrackList.ListItems.Clear
    If Len(Trim(CboPayor)) Then
         GrpCompanyName = Replace(CboPayor, "'", "''")
         GrpCompanyName = Split(GrpCompanyName, "|")(1)
              
         'strAppend = strAppend + " AND HISDB.dbo.MBMOthers.MBMGroupCompany = '" & Trim(GrpCompanyName) & "'"
         strAppend = strAppend + " AND CISWorksheet.dbo.Claim.CLMGrpComCode = '" & Trim(GrpCompanyName) & "'"
    End If
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtCaseNo1)) Then
            strAppend = strAppend + " AND CISDB.dbo.CISCas.CASNumber = '" & Trim(TxtCaseNo1) & "'"
        End If
    Else
        If Len(Trim(TxtCaseNo1)) Then
            strAppend = strAppend + " AND CISDB.dbo.CISCas.CASNumber >= '" & Trim(TxtCaseNo1) & "'"
        End If
        
        If Len(Trim(TxtCaseNo2)) Then
            strAppend = strAppend + " AND CISDB.dbo.CISCas.CASNumber <= '" & Trim(TxtCaseNo2) & "'"
        End If
    End If
    
    If Len(Trim(txtDateVisit1)) > 0 And IsDate(txtDateVisit1) = True Then
        strAppend = strAppend + " And CISCas.DateVisit >= '" & Format(txtDateVisit1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateVisit2)) > 0 And IsDate(txtDateVisit2) = True Then
        strAppend = strAppend + " And CISCas.DateVisit <= '" & Format(txtDateVisit2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkDOV.Value = 1 Then
        sql = ""
        sql = "AND (CISCas.DateVisit Is null OR CISCas.DateVisit = '')" & ""
        strAppend = strAppend + sql
    End If
    If Len(Trim(CboDepartment)) Then
        ''strAppend = strAppend + " AND HISDB.dbo.MBMOthers.MBMDepartment = '" & Trim(Replace(CboDepartment, "'", "''")) & "'"
        strAppend = strAppend + " AND claim.CLMDeptName = '" & Trim(Replace(CboDepartment, "'", "''")) & "'"
    End If
    If Mid(CboClinicType, 1, 1) = 1 Then
        strAppend = strAppend & " AND HISDB.dbo.CLN.CLNType = 'S'"
    ElseIf Mid(CboClinicType, 1, 1) = 2 Then
        strAppend = strAppend & " AND HISDB.dbo.CLN.CLNType = 'G'"
    ElseIf Mid(CboClinicType, 1, 1) = 3 Then
        strAppend = strAppend & " AND  (HISDB.dbo.CLN.CLNType <> 'G') AND (HISDB.dbo.CLN.CLNType <> 'S')"
    End If
    If Len(Trim(txtOPBatchNoFr)) Then
        strAppend = strAppend + " AND CISCas.OPBatchNo >= '" & Trim(txtOPBatchNoFr) & "'"
    End If
    
    If Len(Trim(txtOPBatchNoTo)) Then
        strAppend = strAppend + " AND CISCas.OPBatchNo <= '" & Trim(txtOPBatchNoTo) & "'"
    End If
    
    If Len(Trim(cboSubsidiary)) Then
        ''strAppend = strAppend + " AND HISDB.dbo.MBMOthers.Subsidiary = '" & Trim(Replace(cboSubsidiary, "'", "''")) & "'"
        strAppend = strAppend + " AND  claim.clmsubsidiary = '" & Trim(Replace(cboSubsidiary, "'", "''")) & "'"
    End If
    
    If Len(Trim(cboCashType)) Then
        If Mid(cboCashType, 1, 1) = "C" Then
            strAppend = strAppend + " AND CISWorksheet.dbo.ClaimInvoice.INVPayto = 'C'"
        Else
            strAppend = strAppend + " AND CISWorksheet.dbo.ClaimInvoice.INVPayto = 'M'"
        End If
    End If
    
    If Len(Trim(cboTreatmentCode)) Then
        strAppend = strAppend + " AND CASTreatReason1 = '" & Mid(cboTreatmentCode, 1, 3) & "'"
    End If
    
    If Len(Trim(cboOMNIStatus)) Then
        strAppend = strAppend + " AND GRPCompany.GRPOMNICLM = 'Y'"
    End If
    
    If Len(Trim(txtRegDateFr)) > 0 And IsDate(txtRegDateFr) = True Then
        strAppend = strAppend + " And CISCas.RegisterDate >= '" & Format(txtRegDateFr, "yyyy/MM/DD") & "'"
    End If
    
    If Len(Trim(txtRegDateTo)) > 0 And IsDate(txtRegDateTo) = True Then
        strAppend = strAppend + " And CISCas.RegisterDate <= '" & Format(txtRegDateTo, "yyyy/MM/DD") & "'"
    End If
    If Len(Trim(CboCLN)) Then
        strAppend = strAppend + " AND CISCas.ClinicCode = '" & Trim(Left(CboCLN, IIf(InStr(1, CboCLN, "-") - 1 < 0, Len(CboCLN), InStr(1, CboCLN, "-") - 1))) & "'"
    End If
    
    'If Len(Trim(CboCLN)) Then
    '    strAppend = strAppend + " AND (HISDB.dbo.MBM.HSCLNInd <> 'H') AND " & _
    '                      " (CISWorksheet.dbo.Claim.CLMBordxMemo = 'Y' OR  CISWorksheet.dbo.Claim.CLMBordxMemo = '' OR CISWorksheet.dbo.Claim.CLMBordxMemo IS NULL) AND  " & _
    '                      " (dbo.CISCas.CASStatus <> 'D') "
    'End If
                         
    If cboEntityType.ListIndex = 0 Then
        strAppend = strAppend + " AND Substring(CISCAS.CasNumber, 1, 2) <>'SJ' "
    ElseIf cboEntityType.ListIndex = 1 Then
        strAppend = strAppend + " AND Substring(CISCAS.CasNumber, 1, 2) ='SJ' "
    End If
    
    If ChkScan.Value = 1 Then
        Call WorksheetTrackingListing(True, strAppend)
    Else
        Call WorksheetTrackingListing(False, strAppend)
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Function
'****************** End Search Record *****************************

'****************** Start WorksheetTracking Listing *****************************

Private Sub WorksheetTrackingListing(ScanFlag As Boolean, Optional strAppend As String)
Const METHOD_NAME = "WorksheetTrackingListing"
On Error GoTo debugs

Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim Prm4 As New ADODB.Parameter
Dim strAppend2, tmpbordxtype As String


 'strappend2 = " AND (HISDB.dbo.MBM.HSCLNInd <> 'H') AND (CISWorksheet.dbo.Claim.CLMBordxMemo = 'Y' OR  CISWorksheet.dbo.Claim.CLMBordxMemo = '' OR CISWorksheet.dbo.Claim.CLMBordxMemo IS NULL) AND  (dbo.CISCas.CASStatus <> 'D') "
                      
 'strAppend = strAppend + strappend2
 'IIf(optASO(0).Value = True, 1, 0)
    tmpbordxtype = 0
    If optASO(0).Value = True Then
        tmpbordxtype = 1
    ElseIf optASO(1).Value = True Then
        tmpbordxtype = 0
    ElseIf optASO(2).Value = True Then
        tmpbordxtype = 2
    End If
    total = 0
    Current = 0
    If ScanFlag = True Then
        lstWorksheetTrackList.ListItems.Clear
    End If
    Set cmd.ActiveConnection = MediConnCISWS
    cmd.CommandText = "ClaimBordxModule_ispureaso"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 5000
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, 21)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("BordxType", adTinyInt, adParamInput, 1, CboBordxtype.ListIndex)
    cmd.Parameters.Append Prm3
    Set Prm4 = cmd.CreateParameter("isPureAso", adTinyInt, adParamInput, 1, tmpbordxtype)
    cmd.Parameters.Append Prm4
    
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdBordx.Enabled = True
        lblCurrent = 0
        lblTotal = 0
        LblTotalAmt = 0
    Else
    '# Added by Ravi
    Dim isMultipleCompany As Boolean
    isMultipleCompany = False
    Dim strCompName As String
    '# Added by Ravi
    Do
        Do Until rst2.EOF
        total = rst2.RecordCount
        lblTotal = total
        
            With rst2
                Set Record = lstWorksheetTrackList.ListItems.Add(, , "")
        
                Record.SubItems(1) = IIf(Len(Trim(!INVPayTo)), Trim(!INVPayTo), "")
                Record.SubItems(2) = IIf(Len(Trim(!PAYCode)), Trim(!PAYCode), "")
                Record.SubItems(3) = IIf(Len(Trim(!CasNumber)), Trim(!CasNumber), "")
                Record.SubItems(4) = IIf(IsNumeric(!CLMBordxAmt), Format(!CLMBordxAmt, "#,##0.00"), 0)
                Record.SubItems(5) = IIf(IsNumeric(!ClmNonGPFee), Format(!ClmNonGPFee, "#,##0.00"), 0)
                If Len(!DateVisit) Then
                   Record.SubItems(6) = Format(!DateVisit, "dd/mm/yyyy")
                End If
                Record.SubItems(7) = IIf(Len(Trim(!MBMNumber)), Trim(!MBMNumber), "")
                If !coverid = "00" Then
                    Record.SubItems(8) = IIf(Len(Trim(!MBMName)), Trim(!MBMName), "")
                Else
                    Record.SubItems(8) = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                End If
                Record.SubItems(9) = IIf(Len(Trim(!ClinicCode)), Trim(!ClinicCode), "")
                Record.SubItems(10) = IIf(Len(Trim(!CLNName)), Trim(!CLNName), "")
                Record.SubItems(11) = IIf(Len(Trim(!MBMGroupCompany)), Trim(!MBMGroupCompany), "")
                Record.SubItems(12) = IIf(Len(Trim(!PLNCode)), Trim(!PLNCode), "")
                Record.SubItems(13) = IIf(Len(Trim(!ReasonDesc1)), Trim(!ReasonDesc1), "")
                If !MBMHISPolNo = "03215666" Then
                    Record.SubItems(14) = "AX - AXA AFFIN (NON-EXECUTIVE)"
                ElseIf !MBMHISPolNo = "03215661" Then
                    Record.SubItems(14) = "AX - AXA AFFIN (EXECUTIVE AND ABOVE)"
                Else
                    Record.SubItems(14) = IIf(Len(Trim(!MBMPosition)), Trim(!MBMPosition), "")
                End If
                Record.SubItems(15) = IIf(Len(Trim(!InvoiceNo)), Trim(!InvoiceNo), "")
                Record.SubItems(16) = IIf(Len(Trim(!OPBatchNo)), Trim(!OPBatchNo), "")
                Record.SubItems(17) = IIf(Len(Trim(!CLNFFSStatus)), Trim(!CLNFFSStatus), "")
                If Len(Trim(!CLMMCOProcFee)) Then
                    If InStr(1, !CLMMCOProcFee, "-") Then
                     Record.SubItems(18) = Format(!CLMMCOProcFee, "#,##0.00") * (-1)
                    Else
                     Record.SubItems(18) = Format(!CLMMCOProcFee, "#,##0.00")
                    End If
                Else
                     Record.SubItems(18) = 0
                End If
                Record.SubItems(19) = IIf(Len(Trim(!GRPMCOStatus)), !GRPMCOStatus, "")
                Record.SubItems(20) = IIf(Len(Trim(!CLMTotalActual)), Format(!CLMTotalActual, "#,##0.00"), "")
                Record.SubItems(21) = IIf(Len(Trim(!Subsidiary)), Trim(!Subsidiary), "")
                Record.SubItems(22) = IIf(Len(Trim(!INVBillAmt)), Format(!INVBillAmt, "#,##0.00"), "")
                Record.SubItems(23) = IIf(Len(Trim(!CLMPayableByMedix2c)), Format(!CLMPayableByMedix2c, "#,##0.00"), "")
                Record.SubItems(24) = IIf(Len(Trim(!CLNType)), !CLNType, "")
                Record.SubItems(25) = IIf(Len(Trim(!CASClaimsApp)), !CASClaimsApp, "")
                Record.SubItems(27) = IIf(Len(Trim(!ReasonDesc1)), !ReasonDesc1, "")
                Record.SubItems(28) = IIf(Len(Trim(!ReasonDesc2)), !ReasonDesc2, "")
                Record.SubItems(29) = IIf(Len(Trim(!ReasonDesc3)), !ReasonDesc3, "")
                Record.SubItems(30) = IIf(Len(Trim(!ReasonDesc4)), !ReasonDesc4, "")
                Record.SubItems(31) = IIf(Len(Trim(!ReasonDesc5)), !ReasonDesc5, "")
                Record.SubItems(32) = IIf(Len(Trim(!CASDiagnosis5)), !CASDiagnosis5, "")
                Record.SubItems(33) = IIf(Len(Trim(!MBMPolicyNo)), !MBMPolicyNo, "")
                Record.SubItems(34) = IIf(Len(Trim(!DoctorName)), !DoctorName, "")
                Record.SubItems(35) = IIf(Len(Trim(!CLMCoPayColl)), Format(!CLMCoPayColl, "#,##0.00"), "")
                Record.SubItems(36) = IIf(Len(Trim(!coverid)), Trim(!coverid), "")
                Record.SubItems(37) = IIf(Len(Trim(!clmsubsidiary)), Trim(!clmsubsidiary), "")
                Record.SubItems(38) = IIf(Len(Trim(!clmcostcenter)), Trim(!clmcostcenter), "")
                Record.SubItems(39) = IIf(Len(Trim(!CLMDeptName)), Trim(!CLMDeptName), "")
                
                Current = Current + 1
                lblCurrent = Current
                If Len(!CLMBordxAmt) Then
                    TotalAmt = TotalAmt + !CLMBordxAmt
                End If
'                If Len(!CLMTotalActual) Then
'                    TotalAmt = TotalAmt + !CLMTotalActual
'                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
                '# Added by Ravi
                If isMultipleCompany <> True Then
                    If (strCompName = "") Then
                        strCompName = !MBMGroupCompany
                    ElseIf (!MBMGroupCompany <> strCompName) Then
                        isMultipleCompany = True
                    End If
                End If
              
                If isMultipleCompany = False And strCompName = "UNITED OVERSEAS BANK (MALAYSIA) BHD" Or strCompName = "UOB PROPERTIES BHD" Or strCompName = "UNITED OVERSEAS BANK LIMITED" Then
                    CboUOB.Visible = True
                    lblOPFlex.Visible = True
                    CmdSearch.Enabled = False
                Else
                    CboUOB.Visible = False
                    lblOPFlex.Visible = False
                    CmdSearch.Enabled = True
                End If
                '# Added by Ravi
                
                
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
    CmdExit.Enabled = True
    CmdBordx.Enabled = True
    
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'****************** End WorksheetTracking Listing *****************************

'****************** Start UpdateRecord *****************************

Private Sub UpdateRecord()
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs

Dim strsqlCLM As String
Dim strsqlINV As String
Dim StrSql As String
        
    strsqlCLM = " Update Claim set " & _
                " BordxRefNo = '" & CLMBordxNo & "'," & _
                " ClaimBordxStatus = 1, CLMPending = 'N', " & _
                " CLMBordxMonth = '" & Format(Date, "MM") & "'," & _
                " CLMBordxYear = '" & Format(Date, "yyyy") & "'" & _
                " WHERE CASNumber = '" & Trim(casno) & "'"
    strsqlINV = " update ClaimInvoice set " & _
                " INVBordxRefNo = '" & CLMBordxNo & "'," & _
                " INVBordxed = 1 " & _
                " WHERE CASNumber = '" & Trim(casno) & "'"
    StrSql = " Update CISCAS set " & _
             " Pending = 'N' " & _
             " WHERE CASNumber = '" & Trim(casno) & "'"
    MediConnCISWS.Execute strsqlCLM
    MediConnCISWS.Execute strsqlINV
    MediConnCISDB.Execute StrSql
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub UpdateRecordvoid()
Const METHOD_NAME = "UpdateRecord"
Dim CLNMRec As New ADODB.Recordset
Dim tmpINVPayTo As String
On Error GoTo debugs

Dim strsqlCLM As String
Dim strsqlINV As String
Dim StrSql As String
        
    strsqlCLM = " Update Claim set " & _
                " BordxRefNo = '" & CLMBordxNo & "'," & _
                " ClaimBordxStatus = 1, CLMPending = 'N', " & _
                " CLMBordxMonth = '" & Format(Date, "MM") & "'," & _
                " CLMBordxYear = '" & Format(Date, "yyyy") & "'" & _
                " WHERE CASNumber = '" & Trim(casno) & "'"
    strsqlINV = " update ClaimInvoice set " & _
                " INVBordxRefNo = '" & CLMBordxNo & "'," & _
                " INVBordxed = 1 " & _
                " WHERE CASNumber = '" & Trim(casno) & "'"
    StrSql = " Update CISCAS set " & _
             " Pending = 'N' " & _
             " WHERE CASNumber = '" & Trim(casno) & "'"
    MediConnCISWS.Execute strsqlCLM
    MediConnCISWS.Execute strsqlINV
    MediConnCISDB.Execute StrSql
    
    tmpINVPayTo = ""
    CLNMRec.Open "Select INVPayTo from CISWorksheet..ClaimInvoice WHERE ClaimInvoice.CASNumber ='" & Trim(casno) & "'", MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    If Len(CLNMRec!INVPayTo) Then
         tmpINVPayTo = CLNMRec!INVPayTo
    End If
    CLNMRec.Close
    Set CLNMRec = Nothing
    
    If Trim(tmpINVPayTo) = "M" Then
    StrSql = ""
    StrSql = " INSERT INTO [dbo].[Payment2MBM]([PVNo],[PVDate],[PVChequeDate],[PVChequeNo],[CASNumber],[Pay2MBMLastUpdateDate],[Pay2MBMLastUpdateUser],[PAYAccPeriodMth],[PAYAccPeriodYr],[PVPayMBMStatus])VALUES " & _
                 " ('VOID','" & Format(Date, "yyyy/mm/dd") & "','" & Format(Date, "yyyy/mm/dd") & "','VOID','" & casno & "','" & Format(Date, "yyyy/mm/dd") & "','" & ChkStr(Logon) & "' ,'" & Month(Date) & "','" & Year(Date) & "','F')  "
                  MediConnCISPayment.Execute StrSql
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


'****************** End UpdateRecord *****************************

'****************** Start UpdateBTBGRecord *****************************

Private Sub UpdateBTBGRecord()
Const METHOD_NAME = "UpdateBTBGRecord"
On Error GoTo debugs

Dim strsqlCLM As String
Dim strsqlINV As String
Dim StrSql As String
        
    strsqlCLM = " Update Claim set " & _
                " BTBGBordxRefNo = '" & CLMBordxNo & "'," & _
                " BTBGBordxStatus = 1, CLMPending = 'N' " & _
                " WHERE CASNumber = '" & Trim(casno) & "'"
    strsqlINV = " Update ClaimInvoice set " & _
                " BTBGINVBordxRefNo = '" & CLMBordxNo & "'," & _
                " BTBGINVBordxed = 1 " & _
                " WHERE CASNumber = '" & Trim(casno) & "'"
    StrSql = " Update CISCAS set " & _
             " Pending = 'N' " & _
             " WHERE CASNumber = '" & Trim(casno) & "'"
    MediConnCISWS.Execute strsqlCLM
    MediConnCISWS.Execute strsqlINV
    MediConnCISDB.Execute StrSql
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'****************** End UpdateBTBGRecord *****************************


'****************** Start SaveCLMBordx *****************************

Private Sub SaveCLMBordx(TotalApp As Variant, TotCLM As Integer)
Const METHOD_NAME = "SaveCLMBordx"
On Error GoTo debugs

Dim SqlBordx As String
Dim GrpCoName As String
Dim SubCoName As String
Dim RptType As String
Dim GrpCompanyName As String
    
    GrpCoName = ""
    SubCoName = ""
    
    
     GrpCompanyName = Replace(CboPayor, "'", "''")
         If Len(GrpCompanyName) = 0 Then
            GrpCompanyName = ""
         Else
            GrpCompanyName = Split(GrpCompanyName, "|")(0)
         End If
    GrpCompanyName = Replace(GrpCompanyName, "'", "''")
    SubCoName = Replace(cboSubsidiary, "'", "''")
    If CboReportType.ListIndex = 1 Then
        RptType = 1
    End If
    
    SqlBordx = " Insert Into Bordx(BordxPAYCode, BordxRefNo, " & _
                " BordxDate, BordxBy, BordxClaimAmt, " & _
                " BordxNoClaims, GroupCompany, RptTypeByDept,BordxSubsidary) " & _
                " values('" & Trim(ChkStr(TxtTempPayor)) & "','" & CLMBordxNo & _
                "','" & Format(Date, "yyyy/mm/dd") & "','" & Logon & _
                "'," & TotalApp & "," & TotCLM & ",'" & Trim(GrpCompanyName) & _
                "','" & RptType & "','" & Trim(SubCoName) & "')"
    MediConnCISWS.Execute SqlBordx
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'****************** End SaveCLMBordx *****************************

Private Sub lstWorksheetTrackList_DblClick()
Const METHOD_NAME = "lstWorksheetTrackList_DblClick"
On Error GoTo debugs

    If lstWorksheetTrackList.ListItems.Count > 0 Then
        frmInvoiceAdj.txtCaseNo = Trim(lstWorksheetTrackList.SelectedItem.SubItems(3))
        Call frmInvoiceAdj.CmdGo_Click
        Load frmInvoiceAdj
        frmInvoiceAdj.Show
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

'****************** Start lstWorksheetTrackListItemCheck *****************************
Private Sub lstWorksheetTrackList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Const METHOD_NAME = "lstWorksheetTrackList_ItemCheck"
On Error GoTo debugs

Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Static Arraycnt As Integer
Dim temp
Dim ListCount As Integer
Dim ForLoop As Integer
Dim tmpchkAmt As String
Dim tmpchkDate As String
Dim tmpPatientName As String
Dim tmpchkAmtStatus As Boolean
Dim tmpchkDateStatus As Boolean
Dim tmpPatientNameStatus As Boolean
Dim tmpchkcnt As Integer

    If Item.Checked = True Then
        If Item.SubItems(24) = "G" Then
            tmpchkAmt = Item.SubItems(4)
            tmpchkDate = Item.SubItems(6)
            tmpPatientName = Item.SubItems(8)
            ListCount = 0
            tmpchkcnt = 0
            
            ListCount = lstWorksheetTrackList.ListItems.Count
            For ForLoop = 1 To ListCount
                With lstWorksheetTrackList.ListItems(ForLoop)
                        If lstWorksheetTrackList.ListItems(ForLoop).SubItems(24) = "G" And lstWorksheetTrackList.ListItems(ForLoop).SubItems(25) <> "Y" Then
                            If tmpchkAmt = lstWorksheetTrackList.ListItems(ForLoop).SubItems(4) Then
                                tmpchkAmtStatus = True
                            Else
                                tmpchkAmtStatus = False
                            End If
                            If tmpchkDate = lstWorksheetTrackList.ListItems(ForLoop).SubItems(6) Then
                                tmpchkDateStatus = True
                            Else
                                tmpchkDateStatus = False
                            End If
                            If tmpPatientName = lstWorksheetTrackList.ListItems(ForLoop).SubItems(8) Then
                                tmpPatientNameStatus = True
                            Else
                                tmpPatientNameStatus = False
                            End If
                        End If
                        If tmpchkAmtStatus = True And tmpchkDateStatus = True And tmpPatientNameStatus = True Then
                            tmpchkcnt = tmpchkcnt + 1
                        End If
                        If tmpchkcnt > 1 Then
                            Exit For
                        End If
                End With
            Next

            End If
            If Trim(Item.SubItems(6)) = "" Then
                MsgBox "Date of Visit cannot be empty.", vbCritical
                Item.Checked = False
            ElseIf Trim(Item.SubItems(10)) = "" Then
                MsgBox "Clinic Name cannot be empty!", vbCritical
                Item.Checked = False
            ElseIf Trim(Item.SubItems(11)) = "" Then
                MsgBox "Group Company cannot be empty!", vbCritical
                Item.Checked = False
            ElseIf Trim(Item.SubItems(16)) = "" Then
                MsgBox "Batch Number is empty!", vbCritical
                Item.Checked = False
            ElseIf Trim(Item.SubItems(17)) = "Y" And (Trim(Item.SubItems(18)) = "0" Or Trim(Item.SubItems(18)) = "") Then
                MsgBox "FFS is empty!", vbCritical
                Item.Checked = False
            ElseIf Trim(Item.SubItems(19)) = "" Then
                MsgBox "Group Company's FFS Status is empty!", vbCritical
                Item.Checked = False
            ElseIf Trim(Item.SubItems(4)) = 0 Then
                MsgBox "Please check Bordx Amount!", vbCritical
                Item.Checked = False
            ElseIf Trim(Item.SubItems(1)) = "C" And (Trim(Item.SubItems(20)) <> Trim(Item.SubItems(22))) Then
                MsgBox "Difference in Bill Amount and Payable to Clinic Amount!", vbCritical
                Item.Checked = False
            ElseIf Item.SubItems(35) = 0 And (Trim(Item.SubItems(4)) <> Trim(Item.SubItems(20))) And CheckIsBTBG(Trim(Item.SubItems(12))) = False Then
                MsgBox "Please check Bill Amount & Bordx Amount!", vbCritical
                Item.Checked = False
            Else
                Arraycnt = IIf(UBound(CaseNoArray) = 0, 1, UBound(CaseNoArray) + 1)
                ReDim Preserve CaseNoArray(Arraycnt)
                CaseNoArray(Arraycnt) = Item.SubItems(3)
                PLNCode = Item.SubItems(12)
            End If
    Else
        newbound = 0
        If UBound(CaseNoArray) <> 0 Then
            For i = 0 To UBound(CaseNoArray)
                If (Item.SubItems(3)) <> CaseNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = CaseNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            Erase CaseNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim CaseNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                CaseNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            Arraycnt = 0
        End If
    End If
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If

End Sub
'****************** End lstWorksheetTrackListItemCheck *****************************

Private Sub SetCheckAllItems(bState As Boolean)
Const METHOD_NAME = "SetCheckAllItems"
On Error GoTo debugs
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstWorksheetTrackList.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstWorksheetTrackList.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub SetCheck(hwnd As Long, lItemIndex As Long, bState As Boolean)
Dim LV As LVITEM
Dim r As Long
    With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
    End With
    r = SendMessageAny(hwnd, LVM_SETITEMSTATE, lItemIndex, LV)
End Sub
Private Function CaseNoCheck()
    If Trim(ChkStr(TxtCaseNo1)) = ChkStr(lstWorksheetTrackList.ListItems(listloop1).SubItems(2)) Then
        Check = True
    End If
End Function

'****************** Start ComboListing *****************************
Private Sub PayorComboListing()
Const METHOD_NAME = "PayorComboListing"
On Error GoTo debugs

Dim PAY As New ADODB.Recordset
Dim rst As New ADODB.Recordset
   
    PAY.Open "Select GRPCoName+'|'+GRPCoCode as GRPCoName from GRPCompany ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not PAY.EOF
        CboPayor.AddItem Trim(PAY!GrpCoName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    cboCashType.Clear
    cboCashType.AddItem "C" & "-" & "Cashless"
    cboCashType.AddItem "R" & "-" & "Reimbursement"
    
    cboCashType.Clear
    cboCashType.AddItem "C" & "-" & "Cashless"
    cboCashType.AddItem "R" & "-" & "Reimbursement"
    
    cboOMNIStatus.Clear
    cboOMNIStatus.AddItem "Y" & " - " & "Yes"
    
    
    CboBordxtype.Clear
    CboBordxtype.AddItem "INSURER-KIS" 'KIS
    CboBordxtype.AddItem "ASO-KDC"  'KDC
    CboBordxtype.AddItem "CREDIT NOTE-CN"  'CN
    CboBordxtype.AddItem "SIHAT JOHOR-KSJ"  'SJ
    CboBordxtype.AddItem "VOID-KIV"  'VOID BORDX
    
    cboTreatmentCode.Clear
    rst.Open "Select * from TreatmentReason ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        cboTreatmentCode.AddItem Trim(rst!ReasonCode) & "-" & Trim(rst!ReasonDesc)
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
debugs:
If Err.Description <> "" Then
   Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'****************** End ComboListing *****************************

'****************** Start ComboBox Change/KeyPress*****************************
Private Sub CboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
    TxtCaseNo1.SetFocus
End If
End Sub

Private Sub CboPayor_Click()
    TxtCaseNo1.SetFocus
    '# Added by Ravi
    If (CboPayor.Text <> "UNITED OVERSEAS BANK (MALAYSIA) BHD|U063") Or (CboPayor.Text <> "UOB PROPERTIES BHD|U065") Or (CboPayor.Text <> "UNITED OVERSEAS BANK LIMITED|U066") Then
       lblOPFlex.Visible = False
       CboUOB.Visible = False
    End If
    CmdSearch.Enabled = True
    lstWorksheetTrackList.ListItems.Clear
    '# Added by Ravi
End Sub

Private Sub TxtCaseNo1_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtCaseNo1_LostFocus()
    TxtCaseNo1 = ChkStr(TxtCaseNo1)
    CmdSearch.Default = False
End Sub

'****************** End ComboBox Change/KeyPress*****************************

Private Sub txtDateVisit1_LostFocus()
    If IsDate(txtDateVisit1) Then
    Else
        If txtDateVisit1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateVisit1.SetFocus
        End If
    End If
End Sub

Private Sub txtDateVisit2_LostFocus()
    If IsDate(txtDateVisit2) Then
    Else
        If txtDateVisit2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateVisit2.SetFocus
        End If
    End If
End Sub

Private Sub txtRegDateFr_LostFocus()
    If IsDate(txtRegDateFr) Then
    Else
        If txtRegDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRegDateFr.SetFocus
        End If
    End If

End Sub

Private Sub txtRegDateTo_LostFocus()
    If IsDate(txtRegDateTo) Then
    Else
        If txtRegDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRegDateTo.SetFocus
        End If
    End If

End Sub
