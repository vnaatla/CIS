VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInvAdjustment 
   Caption         =   "Invoice Adjustment"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameCase 
      Height          =   615
      Left            =   4320
      TabIndex        =   147
      Top             =   400
      Width           =   7455
      Begin VB.ComboBox cboPending 
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
         Left            =   6480
         TabIndex        =   4
         Text            =   "N - No"
         Top             =   200
         Width           =   855
      End
      Begin VB.ComboBox cboStatus 
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
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "O - Open"
         Top             =   200
         Width           =   1215
      End
      Begin VB.ComboBox cboClaimability 
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
         Left            =   3880
         TabIndex        =   3
         Text            =   "A - Accepted"
         Top             =   200
         Width           =   1455
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
         Left            =   5640
         TabIndex        =   150
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
         Left            =   240
         TabIndex        =   149
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
         Left            =   2840
         TabIndex        =   148
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   80
      TabIndex        =   141
      Top             =   7560
      Width           =   11655
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10560
         TabIndex        =   68
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9720
         TabIndex        =   67
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdClaerInv 
         Caption         =   "Clear &Inv"
         Height          =   330
         Left            =   7200
         TabIndex        =   64
         Top             =   200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9000
         TabIndex        =   66
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9000
         TabIndex        =   142
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CMdChkItem 
         Caption         =   "C&hk Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         TabIndex        =   65
         Top             =   200
         Width           =   975
      End
      Begin MSComctlLib.ListView lstInvoice 
         Height          =   345
         Left            =   3360
         TabIndex        =   164
         Top             =   240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
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
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   312
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
            Text            =   "Member No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cov ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Patient Name"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "WS Status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Bordxed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Bill No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Stat. Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Visit Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Visit Time"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "MC"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Clinic"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "PreparedBy"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "GP + Med."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "GP + LT Med."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "GP + Med + Inj"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Text            =   "SP+ Med."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Text            =   "SP + LT Med."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   20
            Text            =   "SP + Med + Inj"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   21
            Text            =   "Dental"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   22
            Text            =   "Maternity"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   23
            Text            =   "PRV checkup"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   24
            Text            =   "X-Ray"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   25
            Text            =   "Ultra Sound"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   26
            Text            =   "ECG"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   27
            Text            =   "Dressing"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   28
            Text            =   "Nebulizer"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   29
            Text            =   "I & D"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   30
            Text            =   "T&S/STO"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   31
            Text            =   "Minor Surgery"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "Proc Others"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   33
            Text            =   "Blood Drawing"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   34
            Text            =   "Blood Glucose"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   35
            Text            =   "Urine"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   36
            Text            =   "Pap Smear"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   37
            Text            =   "Lab-Others"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   38
            Text            =   "DTP"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   39
            Text            =   "DTP Polio HIB"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   40
            Text            =   "Hepatitis B"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   41
            Text            =   "MMR"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "Imm Others"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "Co-Payment"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "MedInjType1"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "MedInjType2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "MedInjType3"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Text            =   "MedInjType4"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   48
            Text            =   "MedInjType5"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   49
            Text            =   "MedInjDay1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   50
            Text            =   "MedInjDay2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   51
            Text            =   "MedInjDay3"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   52
            Text            =   "MedInjDay4"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   53
            Text            =   "MedInjDay5"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   54
            Text            =   "MedInjQty1"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   55
            Text            =   "MedInjQty2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   56
            Text            =   "MedInjQty3"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   57
            Text            =   "MedInjQty4"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   58
            Text            =   "MedInjQty5"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   59
            Text            =   "Diagnosis1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   60
            Text            =   "Diagnosis2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   61
            Text            =   "Diagnosis3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   62
            Text            =   "Diagnosis4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   63
            Text            =   "Payor"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   64
            Text            =   "Plan Code"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   65
            Text            =   "Ins Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   66
            Text            =   "HLT Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   67
            Text            =   "Bal Lmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   68
            Text            =   "Visit Lmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   69
            Text            =   "ClnContract"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   70
            Text            =   "Visit IND"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   71
            Text            =   "Bal IND"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   72
            Text            =   "Co-Pay"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   73
            Text            =   "Payto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   74
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   75
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   76
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   77
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   78
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(80) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   79
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(81) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   80
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(82) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   81
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(83) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   82
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(84) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   83
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(85) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   84
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(86) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   85
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(87) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   86
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(88) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   87
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(89) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   88
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(90) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   89
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(91) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   90
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(92) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   91
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(93) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   92
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(94) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   93
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(95) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   94
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(96) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   95
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(97) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   96
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(98) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   97
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(99) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   98
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(100) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   99
            Text            =   "CLMOpenDate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(101) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   100
            Text            =   "CLMOpenBy"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(102) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   101
            Text            =   "CLMUpdateBy"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(103) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   102
            Text            =   "CLMUpdateDate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(104) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   103
            Text            =   "CLMBordxStatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(105) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   104
            Text            =   "CLMStatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(106) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   105
            Text            =   "CLMCloseBy"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(107) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   106
            Text            =   "CLMCloseDate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(108) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   107
            Text            =   "TotFCLN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(109) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   108
            Text            =   "TotCLNExc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(110) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   109
            Text            =   "TotFinalAct"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(111) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   110
            Text            =   "TotFinalAPP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(112) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   111
            Text            =   "TotFinalMem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(113) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   112
            Text            =   "CLMBalLmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(114) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   113
            Text            =   "CLMPrevBal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(115) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   114
            Text            =   "CLMPaidByM2M"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(116) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   115
            Text            =   "CLMPaidByM2C"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(117) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   116
            Text            =   "CLMActGPBMed"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(118) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   117
            Text            =   "CLMActGPMedInj"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(119) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   118
            Text            =   "CLMActGPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(120) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   119
            Text            =   "CLMActSPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(121) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   120
            Text            =   "CLMActSPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(122) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   121
            Text            =   "CLMActSPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(123) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   122
            Text            =   "CLMActXRay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(124) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   123
            Text            =   "CLMActUltraSound"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(125) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   124
            Text            =   "CLMActECG"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(126) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   125
            Text            =   "CLMActDressing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(127) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   126
            Text            =   "CLMActNebulizer"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(128) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   127
            Text            =   "CLMActID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(129) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   128
            Text            =   "CLMActTSSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(130) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   129
            Text            =   "CLMActMinorSurgery"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(131) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   130
            Text            =   "CLMActPOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(132) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   131
            Text            =   "CLMActBloodDrawing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(133) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   132
            Text            =   "CLMActBloodGlucose"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(134) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   133
            Text            =   "CLMActUrine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(135) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   134
            Text            =   "CLMActPapSmear"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(136) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   135
            Text            =   "CLMActLOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(137) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   136
            Text            =   "CLMActDTPPolio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(138) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   137
            Text            =   "CLMActDTPPolioHIB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(139) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   138
            Text            =   "CLMActHepatitisB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(140) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   139
            Text            =   "CLMActMMR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(141) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   140
            Text            =   "CLMActIOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(142) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   141
            Text            =   "CLMActDental"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(143) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   142
            Text            =   "CLMActMaternity"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(144) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   143
            Text            =   "CLMActPRVCheckup"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(145) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   144
            Text            =   "CLMContGPBMed"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(146) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   145
            Text            =   "CLMContGPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(147) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   146
            Text            =   "CLMContGPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(148) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   147
            Text            =   "CLMContSPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(149) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   148
            Text            =   "CLMContSPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(150) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   149
            Text            =   "CLMContSPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(151) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   150
            Text            =   "CLMContXRay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(152) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   151
            Text            =   "CLMContUltraSound"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(153) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   152
            Text            =   "CLMContECG"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(154) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   153
            Text            =   "CLMContDressing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(155) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   154
            Text            =   "CLMContNebulizer"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(156) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   155
            Text            =   "CLMContID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(157) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   156
            Text            =   "CLMContTSSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(158) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   157
            Text            =   "CLMContMinorSurgery"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(159) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   158
            Text            =   "CLMContPOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(160) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   159
            Text            =   "CLMContBloodDrawing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(161) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   160
            Text            =   "CLMContBloodGlucose"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(162) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   161
            Text            =   "CLMContUrine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(163) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   162
            Text            =   "CLMContPapSmear"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(164) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   163
            Text            =   "CLMContLOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(165) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   164
            Text            =   "CLMContDTPPolio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(166) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   165
            Text            =   "CLMContDTPPolioHIB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(167) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   166
            Text            =   "CLMContHepatitisB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(168) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   167
            Text            =   "CLMContMMR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(169) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   168
            Text            =   "CLMContIOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(170) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   169
            Text            =   "CLMContDental"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(171) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   170
            Text            =   "CLMContMaternity"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(172) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   171
            Text            =   "CLMContPRVCheckup"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(173) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   172
            Text            =   "CLMAppGPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(174) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   173
            Text            =   "CLMAppGPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(175) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   174
            Text            =   "CLMAppGPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(176) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   175
            Text            =   "CLMAppSPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(177) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   176
            Text            =   "CLMAppSPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(178) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   177
            Text            =   "CLMAppSPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(179) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   178
            Text            =   "CLMAppXRay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(180) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   179
            Text            =   "CLMAppUltraSound"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(181) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   180
            Text            =   "CLMAppECG"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(182) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   181
            Text            =   "CLMAppDressing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(183) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   182
            Text            =   "CLMAppNebulizer"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(184) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   183
            Text            =   "CLMAppID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(185) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   184
            Text            =   "CLMAppTSSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(186) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   185
            Text            =   "CLMAppMinorSurgery"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(187) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   186
            Text            =   "CLMAppPOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(188) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   187
            Text            =   "CLMAppBloodDrawing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(189) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   188
            Text            =   "CLMAppBloodGlucose"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(190) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   189
            Text            =   "CLMAppUrine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(191) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   190
            Text            =   "CLMAppPapSmear"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(192) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   191
            Text            =   "CLMAppLOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(193) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   192
            Text            =   "CLMAppDTPPolio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(194) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   193
            Text            =   "CLMAppDTPPolioHIB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(195) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   194
            Text            =   "CLMAppHepatitisB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(196) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   195
            Text            =   "CLMAppMMR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(197) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   196
            Text            =   "CLMAppIOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(198) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   197
            Text            =   "CLMAppDental"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(199) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   198
            Text            =   "CLMAppMaternity"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(200) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   199
            Text            =   "CLMAppPRVCheckup"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(201) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   200
            Text            =   "CLMBnfGPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(202) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   201
            Text            =   "CLMBnfGPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(203) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   202
            Text            =   "CLMBnfGPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(204) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   203
            Text            =   "CLMBnfSPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(205) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   204
            Text            =   "CLMBnfSPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(206) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   205
            Text            =   "CLMBnfSPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(207) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   206
            Text            =   "CLMBnfXray"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(208) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   207
            Text            =   "CLMBnfUltraSound"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(209) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   208
            Text            =   "CLMBnfECG"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(210) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   209
            Text            =   "CLMBnfDressing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(211) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   210
            Text            =   "CLMBnfNebulizer"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(212) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   211
            Text            =   "CLMBnfID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(213) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   212
            Text            =   "CLMBnfTSSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(214) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   213
            Text            =   "CLMBnfMinorSurgery"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(215) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   214
            Text            =   "CLMBnfPOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(216) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   215
            Text            =   "CLMBnfBloodDrawing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(217) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   216
            Text            =   "CLMBnfBloodGlucose"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(218) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   217
            Text            =   "CLMBnfUrine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(219) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   218
            Text            =   "CLMBnfPapSmear"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(220) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   219
            Text            =   "CLMBnfLOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(221) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   220
            Text            =   "CLMBnfDTPPolio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(222) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   221
            Text            =   "CLMBnfDTPPolioHIB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(223) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   222
            Text            =   "CLMBnfHepatitisB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(224) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   223
            Text            =   "CLMBnfMMR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(225) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   224
            Text            =   "CLMBnfIOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(226) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   225
            Text            =   "CLMBnfDental"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(227) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   226
            Text            =   "CLMBnfMaternity"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(228) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   227
            Text            =   "CLMBnfPRVCheckup"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(229) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   228
            Text            =   "CLMMBMExcGPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(230) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   229
            Text            =   "CLMMBMExcGPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(231) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   230
            Text            =   "CLMMBMExcGPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(232) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   231
            Text            =   "CLMMBMExcSPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(233) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   232
            Text            =   "CLMMBMExcSPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(234) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   233
            Text            =   "CLMMBMExcSPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(235) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   234
            Text            =   "CLMMBMExcXRay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(236) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   235
            Text            =   "CLMMBMExcUltraSound"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(237) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   236
            Text            =   "CLMMBMExcECG"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(238) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   237
            Text            =   "CLMMBMExcDressing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(239) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   238
            Text            =   "CLMMBMExcNebulizer"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(240) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   239
            Text            =   "CLMMBMExcID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(241) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   240
            Text            =   "CLMMBMExcTSSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(242) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   241
            Text            =   "CLMMBMExcMinorSurgery"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(243) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   242
            Text            =   "CLMMBMExcPOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(244) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   243
            Text            =   "CLMMBMExcBloodDrawing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(245) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   244
            Text            =   "CLMMBMExcBloodGlucose"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(246) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   245
            Text            =   "CLMMBMExcUrine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(247) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   246
            Text            =   "CLMMBMExcPapSmear"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(248) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   247
            Text            =   "CLMMBMExcLOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(249) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   248
            Text            =   "CLMMBMExcDTPPolio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(250) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   249
            Text            =   "CLMMBMExcDTPPolioHIB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(251) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   250
            Text            =   "CLMMBMExcHepatitisB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(252) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   251
            Text            =   "CLMMBMExcMMR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(253) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   252
            Text            =   "CLMMBMExcIOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(254) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   253
            Text            =   "CLMMBMExcDental"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(255) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   254
            Text            =   "CLMMBMExcMaternity"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(256) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   255
            Text            =   "CLMMBMExcPRVCheckup"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(257) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   256
            Text            =   "CLMNPayCLNGPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(258) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   257
            Text            =   "CLMNPayCLNGPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(259) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   258
            Text            =   "CLMNPayCLNGPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(260) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   259
            Text            =   "CLMNPayCLNSPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(261) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   260
            Text            =   "CLMNPayCLNSPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(262) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   261
            Text            =   "CLMNPayCLNSPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(263) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   262
            Text            =   "CLMNPayCLNXRay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(264) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   263
            Text            =   "CLMNPayCLNUltraSound"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(265) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   264
            Text            =   "CLMNPayCLNECG"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(266) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   265
            Text            =   "CLMNPayCLNDressing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(267) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   266
            Text            =   "CLMNPayCLNNebulizer"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(268) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   267
            Text            =   "CLMNPayCLNID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(269) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   268
            Text            =   "CLMNPayCLNTSSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(270) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   269
            Text            =   "CLMNPayCLNMinorSurgery"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(271) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   270
            Text            =   "CLMNPayCLNPThers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(272) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   271
            Text            =   "CLMNPayCLNBloodDrawing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(273) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   272
            Text            =   "CLMNPayCLNBloodGlucose"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(274) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   273
            Text            =   "CLMNPayCLNUrine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(275) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   274
            Text            =   "CLMNPayCLNPapSmear"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(276) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   275
            Text            =   "CLMNPayCLNLOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(277) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   276
            Text            =   "CLMNPayCLNDTPolio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(278) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   277
            Text            =   "CLMNPayCLNDTPPolioHIB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(279) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   278
            Text            =   "CLMNPayCLNHepatitisB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(280) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   279
            Text            =   "CLMNPayCLNMMR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(281) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   280
            Text            =   "CLMNPayCLNIOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(282) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   281
            Text            =   "CLMNPayCLNDental"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(283) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   282
            Text            =   "CLMNPayCLNMaternity"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(284) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   283
            Text            =   "CLMNPayCLNPRVCheckup"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(285) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   284
            Text            =   "CLMPayCLNGPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(286) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   285
            Text            =   "CLMPayCLNGPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(287) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   286
            Text            =   "CLMPayCLNGPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(288) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   287
            Text            =   "CLMPayCLNSPBMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(289) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   288
            Text            =   "CLMPayCLNSPMedInj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(290) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   289
            Text            =   "CLMPayCLNSPLTMed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(291) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   290
            Text            =   "CLMPayCLNXRay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(292) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   291
            Text            =   "CLMPayCLNUltraSound"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(293) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   292
            Text            =   "CLMPayCLNECG"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(294) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   293
            Text            =   "CLMPayCLNDressing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(295) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   294
            Text            =   "CLMPayCLNNebulizer"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(296) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   295
            Text            =   "CLMPayCLNID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(297) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   296
            Text            =   "CLMPayCLNTSSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(298) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   297
            Text            =   "CLMPayCLNMinorSurgery"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(299) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   298
            Text            =   "CLMPayCLNPOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(300) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   299
            Text            =   "CLMPayCLNBloodDrawing"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(301) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   300
            Text            =   "CLMPayCLNBloodGlucose"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(302) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   301
            Text            =   "CLMPayCLNUrine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(303) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   302
            Text            =   "CLMPayCLNPapSmear"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(304) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   303
            Text            =   "CLMPayCLNLOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(305) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   304
            Text            =   "CLMPayCLNDTPPolio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(306) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   305
            Text            =   "CLMPayCLNDTPPolioHIB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(307) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   306
            Text            =   "CLMPayCLNHepatitisB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(308) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   307
            Text            =   "CLMPayCLNMMR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(309) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   308
            Text            =   "CLMPayCLNIOhters"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(310) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   309
            Text            =   "CLMPayCLNDental"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(311) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   310
            Text            =   "CLMPayCLNMaternity"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(312) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   311
            Text            =   "CLMPayCLNPRVCheckup"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label58 
         Caption         =   "Bordx?"
         Height          =   255
         Left            =   360
         TabIndex        =   146
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblBordx 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   145
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label63 
         Caption         =   "WS?"
         Height          =   255
         Left            =   1680
         TabIndex        =   144
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblWs 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   143
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   80
      TabIndex        =   140
      Top             =   400
      Width           =   4215
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
         Left            =   1320
         TabIndex        =   0
         Top             =   195
         Width           =   1575
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   300
         Left            =   3360
         TabIndex        =   1
         Top             =   200
         Width           =   570
      End
      Begin VB.Label Label40 
         Caption         =   "Case Number"
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   240
         Width           =   1095
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
      Height          =   1155
      Left            =   80
      TabIndex        =   124
      Top             =   1080
      Width           =   11775
      Begin VB.TextBox TxtClnContract 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11040
         TabIndex        =   158
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtPolStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9600
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtPolNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
      Begin MSMask.MaskEdBox txtEffDate 
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtExpDate 
         Height          =   315
         Left            =   7560
         TabIndex        =   13
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtMBMStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtTabMBMNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtTabMBMName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtPatientName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtPatientIC 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   11
         Top             =   740
         Width           =   1215
      End
      Begin VB.TextBox txtdob 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtPlanCode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9600
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtInsType 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9600
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox TxtPayor 
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
         Left            =   5280
         TabIndex        =   125
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtVisitLimit 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11040
         TabIndex        =   159
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox TxtBalLimit 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11040
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label54 
         Caption         =   "Visit Lmt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   10200
         TabIndex        =   157
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label34 
         Caption         =   "Cln Contr"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   10200
         TabIndex        =   156
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label64 
         Caption         =   "Plan Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   139
         Top             =   525
         Width           =   735
      End
      Begin VB.Label Label62 
         Caption         =   "Ins Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   138
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label60 
         Caption         =   "Member No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "Pol. Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   135
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Patient DOB"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   134
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Exp. Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   133
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff. Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4480
         TabIndex        =   132
         Top             =   495
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Policy No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4480
         TabIndex        =   131
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Member Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "MBM Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2715
         TabIndex        =   128
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPatientIC 
         Caption         =   "Patient IC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4480
         TabIndex        =   127
         Top             =   765
         Width           =   810
      End
      Begin VB.Label Label5 
         Caption         =   "Bal. Lmt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   10200
         TabIndex        =   126
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label61 
         Caption         =   "Payor"
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
         Left            =   6840
         TabIndex        =   137
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame FrameInvoice 
      Enabled         =   0   'False
      Height          =   5295
      Left            =   80
      TabIndex        =   69
      Top             =   2280
      Width           =   11775
      Begin VB.Frame FrameNotCover 
         Caption         =   "Not Cover"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   6360
         TabIndex        =   116
         Top             =   1840
         Width           =   2250
         Begin VB.TextBox B1 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   51
            Top             =   240
            Width           =   680
         End
         Begin VB.TextBox B2 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   52
            Top             =   480
            Width           =   680
         End
         Begin VB.TextBox B3 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   53
            Top             =   720
            Width           =   680
         End
         Begin VB.TextBox B4 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   54
            Top             =   960
            Width           =   680
         End
         Begin VB.TextBox B5 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   55
            Top             =   1200
            Width           =   680
         End
         Begin VB.TextBox B6 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   56
            Top             =   1440
            Width           =   680
         End
         Begin VB.TextBox B7 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   57
            Top             =   1680
            Width           =   680
         End
         Begin VB.Label Label59 
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
            Left            =   120
            TabIndex        =   123
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label57 
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
            Left            =   120
            TabIndex        =   122
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label56 
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
            Left            =   120
            TabIndex        =   121
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "Maternity"
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
            TabIndex        =   120
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label46 
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
            Left            =   120
            TabIndex        =   119
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label45 
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
            TabIndex        =   118
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label24 
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
            TabIndex        =   117
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame FrameCharges 
         Caption         =   "Charges - Cover"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2200
         Left            =   120
         TabIndex        =   93
         Top             =   1840
         Width           =   6210
         Begin VB.TextBox TxtIndex 
            Height          =   195
            Left            =   960
            TabIndex        =   97
            Top             =   960
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.CommandButton CmdMedication 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Left            =   1960
            TabIndex        =   96
            Top             =   740
            Width           =   240
         End
         Begin VB.CommandButton cmdProcedures 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Left            =   1960
            TabIndex        =   95
            Top             =   980
            Width           =   240
         End
         Begin VB.CommandButton CmdLab 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Left            =   1960
            TabIndex        =   94
            Top             =   1220
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox A1 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   33
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A2 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   34
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A3 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   35
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A4 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   36
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A5 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   37
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox A6 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   38
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A7 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   39
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A8 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   40
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A9 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   41
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A10 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   42
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox A11 
            Alignment       =   1  'Right Justify
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
            Left            =   5475
            MaxLength       =   8
            TabIndex        =   43
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A12 
            Alignment       =   1  'Right Justify
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
            Left            =   5475
            MaxLength       =   8
            TabIndex        =   44
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A13 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   45
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A14 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   46
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A15 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   47
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtOtherDesc1 
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
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   49
            Top             =   1620
            Width           =   2970
         End
         Begin VB.TextBox A16 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   48
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox A17 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   50
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label lblMedicine 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Medication Amount Excess"
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   3440
            TabIndex        =   160
            Top             =   1960
            Visible         =   0   'False
            Width           =   2655
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
            Left            =   4440
            TabIndex        =   115
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label53 
            Caption         =   "Injection"
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
            TabIndex        =   114
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label13 
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
            TabIndex        =   113
            Top             =   1640
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
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   109
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
            Left            =   120
            TabIndex        =   108
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "X-Ray / MRI"
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
            Left            =   2355
            TabIndex        =   107
            Top             =   240
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
            Left            =   2355
            TabIndex        =   106
            Top             =   480
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
            Left            =   2355
            TabIndex        =   105
            Top             =   720
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
            Left            =   2355
            TabIndex        =   104
            Top             =   960
            Width           =   1215
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
            Left            =   2355
            TabIndex        =   103
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Maternity"
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
            TabIndex        =   102
            Top             =   240
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
            Left            =   4440
            TabIndex        =   101
            Top             =   480
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
            Left            =   4440
            TabIndex        =   100
            Top             =   720
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
            Left            =   4440
            TabIndex        =   99
            Top             =   960
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
            Left            =   4440
            TabIndex        =   98
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame FrameCLMCat 
         BackColor       =   &H00C0C0C0&
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
         Height          =   2175
         Left            =   8640
         TabIndex        =   88
         Top             =   1840
         Width           =   3015
         Begin VB.ComboBox CboNatureClm 
            Height          =   315
            Left            =   1320
            TabIndex        =   59
            Top             =   560
            Width           =   1575
         End
         Begin VB.TextBox TxtRemarks 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   60
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtMC 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   58
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblCLnContract 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Excess Amount Basic"
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   120
            TabIndex        =   163
            Top             =   1680
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label lblBalLimit 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Excess Balance Limit"
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   120
            TabIndex        =   162
            Top             =   1470
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label lblVisitLmt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Excess Visit Limit"
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   120
            TabIndex        =   161
            Top             =   1200
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label39 
            Caption         =   "Nature of CLM"
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
            TabIndex        =   154
            Top             =   580
            Width           =   1095
         End
         Begin VB.Label Label41 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Medical Certificate (MC)"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label16 
            Caption         =   "Days"
            Height          =   255
            Left            =   2520
            TabIndex        =   90
            Top             =   270
            Width           =   375
         End
         Begin VB.Label lblTotAmt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Excess Bill"
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   120
            TabIndex        =   89
            Top             =   1875
            Visible         =   0   'False
            Width           =   2775
         End
      End
      Begin VB.Frame FrameRepudiate 
         Caption         =   "Repudiation Reasons"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   84
         Top             =   4080
         Width           =   11505
         Begin VB.ComboBox CboRepudiate1 
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
            Left            =   260
            TabIndex        =   61
            Top             =   240
            Width           =   11160
         End
         Begin VB.ComboBox CboRepudiate2 
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
            Left            =   260
            TabIndex        =   62
            Top             =   480
            Width           =   11160
         End
         Begin VB.ComboBox CboRepudiate3 
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
            Left            =   260
            TabIndex        =   63
            Top             =   720
            Width           =   11160
         End
         Begin VB.Label Label43 
            Caption         =   "1"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   255
            Width           =   255
         End
         Begin VB.Label Label47 
            Caption         =   "2"
            Height          =   180
            Left            =   120
            TabIndex        =   86
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label48 
            Caption         =   "3"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   780
            Width           =   255
         End
      End
      Begin VB.Frame FrameInvDet 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   70
         Top             =   135
         Width           =   11535
         Begin VB.Frame FrameReason 
            Caption         =   "Reason For Seeking Treatment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1220
            Left            =   7320
            TabIndex        =   71
            Top             =   240
            Width           =   4155
            Begin VB.ComboBox CboReason1 
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
               Left            =   240
               TabIndex        =   30
               Tag             =   "0"
               Top             =   220
               Width           =   3840
            End
            Begin VB.ComboBox CboReason2 
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
               Left            =   240
               TabIndex        =   31
               Top             =   500
               Width           =   3840
            End
            Begin VB.ComboBox CboReason3 
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
               Left            =   240
               TabIndex        =   32
               Top             =   780
               Width           =   3840
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
               TabIndex        =   74
               Top             =   240
               Width           =   135
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
               TabIndex        =   73
               Top             =   525
               Width           =   255
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
               TabIndex        =   72
               Top             =   840
               Width           =   255
            End
         End
         Begin VB.TextBox txtInvoiceNo 
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
            Left            =   1320
            TabIndex        =   21
            Top             =   470
            Width           =   1455
         End
         Begin VB.OptionButton OptMultiple 
            Caption         =   "Multiple"
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
            TabIndex        =   20
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton OptChronic 
            Caption         =   "Chronic"
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
            Left            =   1200
            TabIndex        =   19
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton OptAccute 
            Caption         =   "Accute"
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
            TabIndex        =   18
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtInvoiceDate 
            Height          =   285
            Left            =   1320
            TabIndex        =   22
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox txtDateVisit 
            Height          =   285
            Left            =   1320
            TabIndex        =   23
            Top             =   970
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
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
         Begin VB.ComboBox CboPAYto 
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
            Left            =   3600
            TabIndex        =   25
            Text            =   "C - Clinic"
            Top             =   420
            Width           =   1335
         End
         Begin VB.TextBox TxtInvPayType 
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
            Height          =   315
            Left            =   6040
            TabIndex        =   26
            Text            =   "Cashless"
            Top             =   420
            Width           =   1095
         End
         Begin VB.ComboBox Cboclinic 
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
            Left            =   3600
            TabIndex        =   27
            Top             =   705
            Width           =   3540
         End
         Begin MSMask.MaskEdBox txtVisitTime 
            Height          =   315
            Left            =   1320
            TabIndex        =   24
            Top             =   1220
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtPayee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   28
            Top             =   960
            Width           =   3540
         End
         Begin VB.TextBox txtDoctor 
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
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1270
            Width           =   3540
         End
         Begin VB.Label Label66 
            Caption         =   "Visit Time"
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
            TabIndex        =   155
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label LblExpDate 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Expired"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   5640
            TabIndex        =   153
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Submission Date"
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
            TabIndex        =   83
            Top             =   1000
            Width           =   1215
         End
         Begin VB.Label Label14 
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
            Left            =   2925
            TabIndex        =   82
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Invoice No"
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
            TabIndex        =   81
            Top             =   500
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Invoice Date"
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
            TabIndex        =   80
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Pay To"
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
            Left            =   2925
            TabIndex        =   79
            Top             =   435
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Payee"
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
            Left            =   2925
            TabIndex        =   78
            Top             =   1035
            Width           =   975
         End
         Begin VB.Label Label11 
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
            Left            =   5115
            TabIndex        =   77
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Lbl45days 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "45 Days Late"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   3600
            TabIndex        =   76
            Top             =   165
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label50 
            Caption         =   "Doctor"
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
            Left            =   2925
            TabIndex        =   75
            Top             =   1320
            Width           =   735
         End
      End
   End
   Begin VB.Label Label35 
      BackColor       =   &H00404040&
      Caption         =   "  Invoice Adjustment"
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
      TabIndex        =   151
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmInvAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bExit As Boolean
Dim intIndex As String
Dim EditFlag As Boolean
'Dim objword As Word.Application
'Dim objdoc As Word.Document
Dim ScanCode As String
Dim strBillChkItem, strMEdChkItem As String
Dim strBillChkAmt, strMEdChkAmt, IntInvAmt As Integer
Dim CASENumber, NewCaseNo, MembershipNo As String
Dim strPolNo, strMBMName, strPatientName, strMBMDataStatus, strPOlStatus As String
Dim strEffdate, strExpDate, strPlanCode, strInsType, strPatientDOB As String
Dim strPayor, strPatientID, strPatientIC, strBalLimit, strVisitLimit, strCLNContract As String
Private Hospitals As String

Private Sub A1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A3_GotFocus()
    lblMedicine.Visible = False
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
Private Sub B1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B7_KeyPress(KeyAscii As Integer)
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


Private Sub CboNatureClm_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboNatureClm.Text, CboNatureClm.SelStart) + Chr(KeyAscii) + Mid(CboNatureClm.Text, CboNatureClm.SelStart + CboNatureClm.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboNatureClm.ListCount - 1
       
        If NewText = LCase(Left(CboNatureClm.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboNatureClm.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboNatureClm.Text = ValidValue
            CboNatureClm.SelStart = 0
            CboNatureClm.SelLength = Len(ValidValue)
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

Private Sub CboRepudiate1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRepudiate1.Text, CboRepudiate1.SelStart) + Chr(KeyAscii) + Mid(CboRepudiate1.Text, CboRepudiate1.SelStart + CboRepudiate1.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboRepudiate1.ListCount - 1
       
        If NewText = LCase(Left(CboRepudiate1.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboRepudiate1.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboRepudiate1.Text = ValidValue
            CboRepudiate1.SelStart = 0
            CboRepudiate1.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboRepudiate2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRepudiate2.Text, CboRepudiate2.SelStart) + Chr(KeyAscii) + Mid(CboRepudiate2.Text, CboRepudiate2.SelStart + CboRepudiate2.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboRepudiate2.ListCount - 1
       
        If NewText = LCase(Left(CboRepudiate2.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboRepudiate2.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboRepudiate2.Text = ValidValue
            CboRepudiate2.SelStart = 0
            CboRepudiate2.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboRepudiate3_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRepudiate3.Text, CboRepudiate3.SelStart) + Chr(KeyAscii) + Mid(CboRepudiate2.Text, CboRepudiate3.SelStart + CboRepudiate3.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboRepudiate3.ListCount - 1
       
        If NewText = LCase(Left(CboRepudiate3.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboRepudiate3.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboRepudiate3.Text = ValidValue
            CboRepudiate3.SelStart = 0
            CboRepudiate3.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CMdChkItem_Click()
    Call CheckItemsAmt
End Sub

Private Sub CmdClaerInv_Click()
    Call ClearInvoice
End Sub


Private Sub CmdMedication_Click()
    frmINVMedicine.Source = 1
    frmINVMedicine.Show
End Sub

Private Sub cmdProcedures_Click()
    frmINVMedicine.Source = 2
    frmINVMedicine.Show
End Sub

Private Sub CmdLab_Click()
'    frmINVMedicine.Source = 3
'    frmINVMedicine.Show
End Sub

Public Sub cmdShow_Click()
    CASENumber = ChkStr(txtCaseNo)
    CmdEdit.Enabled = True
    CmdEdit.Visible = True
    CmdUpdate.Visible = False
    Call InvoiceINFO
    Call MBMInfo
    Call CheckItemsAmt
    If Trim(MembershipNo) <> "" Then
        txtTabMBMNo = MembershipNo
        TxtPayor = strPayor
        txtPlanCode = strPlanCode
        TxtInsType = strInsType
        txtMBMName = strMBMName
        txtTabMBMName = strMBMName
        txtEffDate = strEffdate
        txtExpDate = strExpDate
        TxtPolNo = strPolNo
        txtPolStatus = strPOlStatus
        txtMBMStatus = strMBMDataStatus
        txtdob = strPatientDOB
        txtPatientIC = strPatientIC
        txtPatientName = strPatientName
        TxtBalLimit = strBalLimit
        TxtVisitLimit = strVisitLimit
        TxtClnContract = strCLNContract
    End If
End Sub


Private Sub OptAccute_LostFocus()
    If OptAccute.Value = True Then
        strBillChkItem = "BILL ACCUTE"
        strMEdChkItem = "MEDICINE ACCUTE"
    End If
End Sub

Private Sub OptChronic_LostFocus()
    If OptChronic.Value = True Then
        strBillChkItem = "BILL CHRONIC"
        strMEdChkItem = "MEDICINE CHRONIC"
    End If
End Sub

Private Sub OptMultiple_LostFocus()
    If OptMultiple.Value = True Then
        strBillChkItem = "BILL MULTIPLE"
        strMEdChkItem = "MEDICINE MULTIPLE"
    End If
End Sub

Private Sub txtCaseNo_GotFocus()
    cmdShow.Default = True
End Sub

Private Sub txtCaseNo_LostFocus()
    If Trim(txtCaseNo) <> "" Then
        CASENumber = txtCaseNo
    End If
    cmdShow.Default = False
End Sub

Private Sub txtDoctor_LostFocus()
    txtDoctor = ChkStr(txtDoctor)
End Sub

Private Sub txtMC_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Cboclinic_LostFocus()
    If Mid(CboPAYto, 1, 1) = "C" Then
        TxtPayee = Cboclinic
    End If
End Sub

Private Sub CboPAYto_LostFocus()
    If Mid(CboPAYto, 1, 1) = "C" Then
        TxtInvPayType = "Cashless"
    Else
        TxtInvPayType = "Reimbursement"
        txtpayeet = "MEMBER"
    End If
End Sub



Private Sub CmdClear_Click()
    Call ClearEntries
    strMBMno = ""
    strPayor = ""
    strPlanCode = ""
    strInsType = ""
    strMBMName = ""
    strTabMBMName = ""
    strEffdate = ""
    strExpDate = ""
    strPolNo = ""
    strMBMDataStatus = ""
    strPOlStatus = ""
    strPatientName = ""
    strPatientDOB = ""
    strPatientIC = ""
    strPatientID = ""
    strBalLimit = ""
    strVisitLimit = ""
    strCLNContract = ""
    TxtInvPayType = "Cashless"
    OptAccute.Value = True
    strBillChkItem = "BILL ACCUTE"
    strMEdChkItem = "MEDICINE ACCUTE"
    MembershipNo = ""
    CASENumber = ""
    lblTotAmt.Visible = False
    lblMedicine.Visible = False
    LblExpDate.Visible = False
    lblVisitLmt.Visible = False
    lblBalLimit.Visible = False
    lblCLnContract.Visible = False
    Unload frmINVMedicine
    Call EntriesEnable(False)
End Sub

Private Sub cmdEdit_Click()
    CmdUpdate.Visible = True
    CmdUpdate.Enabled = True
    CmdUpdate.Caption = "&Update"
    CmdEdit.Visible = False
    CmdEdit.Enabled = False
    Call EntriesEnable(True)
    EditFlag = True
    txtInvoiceNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdUpdate_Click()
Dim CISCasRec As New ADODB.Recordset
Dim CISCasSql As String
Dim InvChk As Boolean
Dim DupInvoice

    If txtInvoiceDate = "__/__/____" Or txtInvoiceDate = "" Then
        MsgBox "Please key in Invoice date! ", vbCritical
        txtInvoiceDate.SetFocus
    ElseIf txtDateVisit = "__/__/____" Or txtDateVisit = "" Then
        MsgBox "Please Enter Submission Date! ", vbCritical
        txtDateVisit.SetFocus
     ElseIf Trim(Cboclinic) = "" Then
        MsgBox "Please keyin the Clinic code ! ", vbCritical
        Cboclinic.SetFocus
    ElseIf Trim(CboPAYto) = "" Then
        MsgBox "Please select who to be paid (Member / Clinic )! ", vbCritical
        CboPAYto.SetFocus
    ElseIf Trim(TxtPayee) = "" Then
        MsgBox "Please keyin the person/Clinic to be paid! ", vbCritical
        TxtPayee.SetFocus
    ElseIf Mid(CboPAYto, 1, 1) = "C" And Trim(Left(Cboclinic, IIf(InStr(1, Cboclinic, "-") - 1 < 0, Len(Cboclinic), InStr(1, Cboclinic, "-") - 1))) <> Trim(Left(TxtPayee, IIf(InStr(1, TxtPayee, "-") - 1 < 0, Len(TxtPayee), InStr(1, TxtPayee, "-") - 1))) Then
        MsgBox "Clinic code is different with the Payee code ", vbCritical
        Cboclinic.SetFocus
    Else
        CISCasSql = "SELECT * FROM View_Clm_Invoice where CASNumber = '" & Trim(CASENumber) & "'"
        CISCasRec.Open CISCasSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        If CISCasRec.RecordCount And EditFlag = False Then
            If InvChk = False Then
                DupInvoice = MsgBox("Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
                If DupInvoice = vbYes Then
                    Call SaveRecord
                ElseIf DupInvoice = vbNo Then
                    Call EntriesEnable(False)
                End If
            Else
                MsgBox "Record Duplicated, Same Invoice has been entered !", vbCritical, "Error"
            End If
            
            Exit Sub
        Else
            Call SaveRecord
        End If
        CmdUpdate.Visible = False
        CISCasRec.Close
        Set CISCasRec = Nothing
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
    bExit = False
    EditFlag = False
    Call EntriesEnable(False)
    Call ComboListing
    CmdClear.Enabled = True
    CmdUpdate.Visible = False
    strBillChkItem = "BILL ACCUTE"
    strMEdChkItem = "MEDICINE ACCUTE"
    intIndex = 0
End Sub

Private Sub ComboListing()
Dim ClinicRec As New ADODB.Recordset
Dim RepRec As New ADODB.Recordset
Dim TreatRec As New ADODB.Recordset
Dim NatureCLM As New ADODB.Recordset
        
    TxtInvPayType = "Cashless"
    Cboclinic.Clear
 '   CboPayee.Clear
    ClinicRec.Open "SELECT CLNCode, CLNName FROM CLN", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until ClinicRec.EOF
        Cboclinic.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
'        CboPayee.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
        ClinicRec.MoveNext
    Loop
    ClinicRec.Close
    Set ClinicRec = Nothing
    
    CboRepudiate1.Clear
    CboRepudiate2.Clear
    CboRepudiate3.Clear
    RepRec.Open "SELECT REPCode, RepDescription FROM REPCode", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until RepRec.EOF
        CboRepudiate1.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        CboRepudiate2.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        CboRepudiate3.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        RepRec.MoveNext
    Loop
    RepRec.Close
    Set RepRec = Nothing
    
    CboReason1.Clear
    CboReason2.Clear
    CboReason3.Clear
    TreatRec.Open "SELECT ReasonCode, ReasonDesc FROM TreatmentReason", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until TreatRec.EOF
        CboReason1.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        CboReason2.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        CboReason3.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        TreatRec.MoveNext
    Loop
    TreatRec.Close
    Set RepRec = Nothing
    
    CboNatureClm.Clear
    NatureCLM.Open "SELECT ClaimNatureCode, ClaimNatureName FROM ClaimNature", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until NatureCLM.EOF
        CboNatureClm.AddItem NatureCLM!ClaimNatureCode & " - " & NatureCLM!ClaimNatureName
        NatureCLM.MoveNext
    Loop
    NatureCLM.Close
    Set NatureCLM = Nothing
    
    CboPAYto.Clear
    CboPAYto.AddItem "C - Clinic"
    CboPAYto.AddItem "M - Member"
    CboPAYto.ListIndex = 0

    cboClaimability.Clear
    cboClaimability.AddItem "A - Accepted"
    cboClaimability.AddItem "R - Rejected"
    cboClaimability.AddItem "C - Cancelled"
    cboClaimability.ListIndex = 0
    
    cboPending.Clear
    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"
    cboPending.ListIndex = 1
    
    cboStatus.Clear
    cboStatus.AddItem "O - Open"
    cboStatus.AddItem "C - Close"
    cboStatus.AddItem "D - Delete"
    cboStatus.ListIndex = 0
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
    txtVisitTime.mask = ""
    txtVisitTime.Text = ""
    txtVisitTime.mask = "##:##"
    TxtInvPayType = "Cashless"
    txtINvDate = ""
    Cboclinic.Text = ""
    TxtPayee.Text = ""
    CboNatureClm.Text = ""
    CboReason1.Text = ""
    CboReason2.Text = ""
    CboReason3.Text = ""
    CboRepudiate1.Text = ""
    CboRepudiate2.Text = ""
    CboRepudiate3.Text = ""
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
    lblBordx = ""
    lblVersion = ""
    lblLastVersion = ""
    Lbl45days.Visible = False
    lblWs = ""
    IntInvAmt = 0
End Sub

Private Sub txtInvoiceNo_LostFocus()
    txtInvoiceNo = ChkStr(txtInvoiceNo)
End Sub

Private Sub TxtPayee_LostFocus()
    TxtPayee = ChkStr(TxtPayee)
End Sub

Private Sub txtVisitTime_LostFocus()
    If txtVisitTime <> "__:__" Then
        If isTime(txtVisitTime) = False Then
            MsgBox "Invalid Time Format! ", vbCritical
            txtVisitTime.SetFocus
        End If
    End If
End Sub

Private Sub txtDateVisit_LostFocus()
    If txtDateVisit <> "__/__/____" Then
        If IsDate(txtDateVisit) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtDateVisit.SetFocus
        ElseIf CDate(txtDateVisit) > CDate(Date) Then
            MsgBox "Submission Date cannot greater than system date ! ", vbCritical
            txtDateVisit.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceDate_LostFocus()
    If txtInvoiceDate <> "__/__/____" Then
        If IsDate(txtInvoiceDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInvoiceDate.SetFocus
        ElseIf CDate(txtInvoiceDate) > CDate(Date) Then
            MsgBox "Invoice date cannot greater than system date ! ", vbCritical
            txtInvoiceDate.SetFocus
        ElseIf CDate(txtInvoiceDate) < CDate(strEffdate) Then
            MsgBox "Invoice date must be greater than Effective date ! ", vbCritical
            txtInvoiceDate.SetFocus
        ElseIf CDate(txtInvoiceDate) > CDate(strExpDate) + 14 Then
            MsgBox "Invoice date cannot greater than Expiry date + 14 days ! ", vbCritical
            txtInvoiceDate.SetFocus
        Else
            If CDate(Date) - CDate(txtInvoiceDate) > 45 Then
                Lbl45days.Visible = True
            Else
                Lbl45days.Visible = False
            End If
        '    txtDateVisit = txtInvoiceDate
        End If
    End If
End Sub

Private Sub MBMInfo()
Dim MBMRec As New ADODB.Recordset
Dim CLNOth As New ADODB.Recordset
Dim MBMSql, CLNSql, COVSql, CISAnnLmtSql As String
Dim COVRec As New ADODB.Recordset
Dim Cnt As Integer
Dim CISPln As New ADODB.Recordset
Dim CISAnnLMt As New ADODB.Recordset
Dim COVLmt As New ADODB.Recordset
Dim COVLmtSql As String

    If Len(Trim(MembershipNo)) Then
        MBMSql = "SELECT MBMNumber, MBMName, MBMICBcPp, PAYCode, MBMCLNBalLmt, MBMCLNVisitLmt, " & _
                " MBMStatus, MBMCOVID, MBMDOB, MBMDataStatus, HLTCode FROM MBM WHERE MBMNumber='" & Trim(MembershipNo) & " ' AND MBMClinical='Y' AND MBMDataStatus <> 'D' AND MBMStatus <> 'D'"
        MBMRec.Open MBMSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        CLNSql = "SELECT MBMNumber, MBMPolNumber, MBMPolEffDate, MBMPolExpDate, " & _
                " INSType, PLNType FROM MBMCLNOthers WHERE MBMNumber='" & Trim(MembershipNo) & "'"
        CLNOth.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        CISPlnSql = "SELECT AnnLmtIND, VisLmtIND, ClnContract from PLN WHERE PLNCode ='" & Trim(CLNOth!PLNType) & "' and HLTCode='" & Trim(MBMRec!HLTCode) & "'"
        CISPln.Open CISPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If Trim(CISPln!AnnLmtIND) = "Y" Then
            CISAnnLmtSql = "SELECT SuppLimitStatus FROM AnnualLimit WHERE PLNCode='" & Trim(CLNOth!PLNType) & "' and HLTCode='" & Trim(MBMRec!HLTCode) & "' AND INSCode='" & Trim(CLNOth!INSType) & "'"
            CISAnnLMt.Open CISAnnLmtSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            SuppLmtType = CISAnnLMt!SuppLimitStatus
            CISAnnLMt.Close
            Set CISAnnLMt = Nothing
        Else
            SuppLmtType = ""
        End If
        If Trim(CISPln!AnnLmtIND) = "N" Then
            strBalLimit = "FULL"
        Else
            strBalLimit = MBMRec!MBMCLNBalLmt
        End If
        If Trim(CISPln!VisLmtIND) = "N" Then
            strVisitLimit = "FULL"
        Else
            strVisitLimit = MBMRec!MBMCLNVisitLmt
        End If
        strCLNContract = CISPln!ClnContract
        CISPln.Close
        Set CISPln = Nothing
        
        strMBMno = MBMRec!MBMNumber
        strPayor = MBMRec!PAYCode
        strPlanCode = CLNOth!PLNType
        strInsType = CLNOth!INSType
        strMBMName = MBMRec!MBMName
        strTabMBMName = MBMRec!MBMName
        strEffdate = CLNOth!MBMPolEffDate
        strExpDate = CLNOth!MBMPolExpDate
        strPolNo = CLNOth!MBMPolNumber
        strMBMDataStatus = MBMRec!MBMDataStatus
        strPOlStatus = MBMRec!MBMStatus
        strPatientName = MBMRec!MBMName
        strPatientDOB = MBMRec!MBMDOB
        strPatientIC = MBMRec!MBMIcBcPp

        COVSql = "SELECT MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB, MBMCAnnualLimit, MBMCAvailableLimit FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(strPatientID) & "'"
        COVRec.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If COVRec.RecordCount Then
            strPatientName = COVRec!MBMCName
            strPatientDOB = COVRec!MBMCDOB
            strPatientIC = COVRec!MBMCICBCPPNo
            strPatientID = COVRec!MBMCCoverID
            If SuppLimitStatus = "B" Then
                If CovPersonID <> "01" Then
                    COVLmtSql = "SELECT MBMCCLNBalLmt, MBMCCLNVisitLmt FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '01'"
                    COVLmt.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    strBalLimit = COVLmt!MBMCCLNBalLmt
                    strVisitLimit = COVLmt!MBMCCLNVisitLmt
                    COVLmt.Close
                    Set COVLmt = Nothing
                End If
            ElseIf SuppLimitStatus = "C" Then
                COVLmtSql = "SELECT MBMCCLNBalLmt, MBMCCLNVisitLmt FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(CovPersonID) & "'"
                COVLmt.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                strBalLimit = COVLmt!MBMCCLNBalLmt
                strVisitLimit = COVLmt!MBMCCLNVisitLmt
                COVLmt.Close
                Set COVLmt = Nothing
            End If
        End If
        COVRec.Close
        Set COVRec = Nothing
        If Date - CDate(strExpDate) > 14 Then
            LblExpDate.Visible = True
        Else
            LblExpDate.Visible = False
        End If
        MBMRec.Close
        Set MBMRec = Nothing
        CLNOth.Close
        Set CLNOth = Nothing
    End If
    
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
    CASENumber = txtCaseNo
    strsql = " WHERE CASNumber='" & Trim(CASENumber) & "'"
    CASSql = "SELECT MBMNumber, * FROM View_Clm_Invoice " & strsql
    CISCAS1.Open CASSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    If CISCAS1.RecordCount <= 0 Then
        MsgBox "Not Record Found ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Call ClearEntries
        txtCaseNo = CASENumber
        MembershipNo = ""
        CASENumber = ""
        txtCaseNo.SetFocus
    Else
        With CISCAS1
            MembershipNo = !MBMNumber
            CASENumber = !CASNumber
            strPatientID = !COVERID
            TxtIndex = !InvINdex
            txtInvoiceNo = !INVNo
            txtInvoiceDate = !InvDate
            txtDateVisit = !DateVisit
            If Len(Trim(!TimeVisit)) Then
                txtVisitTime = Format(!TimeVisit, "hh:mm")
            End If
            txtDoctor = !CASDoctor
            Cboclinic = IIf(Len(!ClinicCode), !ClinicCode, "")
            CboNatureClm = IIf(Len(!NatureCLM), !NatureCLM, "")
            OptAccute.Value = False
            OptChronic.Value = False
            OptChronic.Value = False
            If !ChkItemCat = "BILL ACCUTE" Then
                OptAccute.Value = True
            ElseIf !ChkItemCat = "BILL CHRONIC" Then
                OptChronic.Value = True
            ElseIf !ChkItemCat = "BILL MULTIPLE" Then
                OptChronic.Value = True
            End If
            TxtRemarks = IIf(Len(!Remarks), !Remarks, "")
            CboPAYto = !INVPayTo
            If Mid(CboPAYto, 1, 2) = "C" Then
                TxtInvPayType = "Cashless"
            Else
                TxtInvPayType = "Reimbursement"
            End If
            TxtPayee = !INVPayee
            A1 = IIf(Len(!CLMGPConsAmt), !CLMGPConsAmt, 0)
            A2 = IIf(Len(!CLMSPConsAmt), !CLMSPConsAmt, 0)
            A3 = IIf(Len(!CLMMedicine), !CLMMedicine, 0)
            A4 = IIf(Len(!CLMProcedures), !CLMProcedures, 0)
            A5 = IIf(Len(!CLMLab), !CLMLab, 0)
            A6 = IIf(Len(!CLMXray), !CLMXray, 0)
            A7 = IIf(Len(!CLMDental), !CLMDental, 0)
            A8 = IIf(Len(!CLMChineseMed), !CLMChineseMed, 0)
            A9 = IIf(Len(!CLMTraditionalMEd), !CLMTraditionalMEd, 0)
            A10 = IIf(Len(!CLMphysioTheraphy), !CLMphysioTheraphy, 0)
            A11 = IIf(Len(!CLMAccident), !CLMAccident, 0)
            A12 = IIf(Len(!CLMPRVCheckup), !CLMPRVCheckup, 0)
            A13 = IIf(Len(!CLMVision), !CLMVision, 0)
            A14 = IIf(Len(!CLMID), !CLMID, 0)
            A15 = IIf(Len(!CLMTSSTO), !CLMTSSTO, 0)
            A16 = IIf(Len(!CLMInjection), !CLMInjection, 0)
            A17 = IIf(Len(!CLMOTHERS1), !CLMOTHERS1, 0)
            
            B1 = IIf(Len(!SPFee), !SPFee, 0)
            B2 = IIf(Len(!Medicine), !Medicine, 0)
            B3 = IIf(Len(!Dental), !Dental, 0)
            B4 = IIf(Len(!Physio), !Physio, 0)
            B5 = IIf(Len(!Maternity), !Maternity, 0)
            B6 = IIf(Len(!PRVCheckup), !PRVCheckup, 0)
            B7 = IIf(Len(!Vision), !Vision, 0)
            
            If !Claimability = "A" Then
                cboClaimability.Text = "A - Accepted"
            ElseIf !Claimability = "R" Then
                cboClaimability.Text = "R - Rejected"
            ElseIf !Claimability = "C" Then
                cboClaimability.Text = "C - Cancelled"
            End If
            If !Pending = "Y" Then
                cboPending.Text = "Y - Yes"
            ElseIf !Pending = "N" Then
                cboPending.Text = "N - No"
            End If
            CboReason1.Text = IIf(Len(!CASTreatReason1), !CASTreatReason1, "")
            CboReason2.Text = IIf(Len(!CASTreatReason2), !CASTreatReason2, "")
            CboReason3.Text = IIf(Len(!CASTreatReason3), !CASTreatReason3, "")
            txtMC = !MCDays
            CboRepudiate1.Text = IIf(Len(!CASREPCode1), !CASREPCode1, "")
            CboRepudiate2.Text = IIf(Len(!CASREPCode2), !CASREPCode2, "")
            CboRepudiate3.Text = IIf(Len(!CASREPCode3), !CASREPCode3, "")
            lblBordx = !INVBordxed
            lblWs = !INVWSStatus
            
'  frmInvMedicine
            frmINVMedicine.txtMEdication1.Text = IIf(Len(!CLMMedication1), !CLMMedication1, "")
            frmINVMedicine.txtMEdication2.Text = IIf(Len(!CLMMedication2), !CLMMedication2, "")
            frmINVMedicine.txtMEdication3.Text = IIf(Len(!CLMMedication3), !CLMMedication3, "")
            frmINVMedicine.CboProc1.Text = IIf(Len(!CLMProc1), !CLMProc1, "")
            frmINVMedicine.CboProc2.Text = IIf(Len(!CLMProc2), !CLMProc2, "")
            frmINVMedicine.CboProc3.Text = IIf(Len(!CLMProc3), !CLMProc3, "")
            frmINVMedicine.CboProc4.Text = IIf(Len(!CLMProc4), !CLMProc4, "")
            frmINVMedicine.CboProc5.Text = IIf(Len(!CLMProc5), !CLMProc5, "")
            frmINVMedicine.CboProc6.Text = IIf(Len(!CLMProc6), !CLMProc6, "")
            frmINVMedicine.CboProc7.Text = IIf(Len(!CLMProc7), !CLMProc7, "")
            
        End With
        Call EntriesEnable(False)
    End If
    CISCAS1.Close
    Set CISCAS1 = Nothing
End Sub

Private Sub SaveRecord()
Dim RecordSaved
Dim startTrans As Boolean
    
    Call CheckItemsAmt
    RecordSaved = MsgBox("Are you sure ? ", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
            MediConnCISWS.beginTrans
            startTrans = True
               
            Call getTotal
            Call SaveCase
            Call SaveInvoice
            Call saveActual
            Call saveActualOthers
'            Call SaveReason
            Call saveNotCover
        
            MediConnCISWS.CommitTrans
            startTrans = False
            MsgBox "Record Updated ! ", vbOKOnly
            
            CmdClear.Enabled = True
'            Call ClearEntries
            Call InvoiceINFO
            Call MBMInfo
            Call EntriesEnable(False)
            CmdEdit.Visible = True
            CmdEdit.Enabled = True
            EditFlag = False
    End If
    Screen.MousePointer = vbDefault
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
    IntInvAmt = 0
    For Each ctl In Me.Controls
        If Mid(ctl.Name, 1, 1) = "A" And Len(ctl.Name) <= 7 Then
            If IsNumeric(ctl.Text) = True Then
                IntInvAmt = IntInvAmt + CDbl(ctl.Text)
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
        !MCDays = IIf(IsNumeric(txtMC), txtMC, 0)
        !INVLastUpdateBy = Logon
        !INVLastUpdateDate = Date
        If IsNumeric(IntInvAmt) = True Then
            !INVAmount = IntInvAmt
        Else
            !INVAmount = 0
        End If
        !INVPayee = Trim(Left(TxtPayee, IIf(InStr(1, TxtPayee, "-") - 1 < 0, Len(TxtPayee), InStr(1, TxtPayee, "-") - 1)))
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
        !CLMDental = IIf(IsNumeric(A7), A7, 0)
        !CLMChineseMed = IIf(IsNumeric(A8), A8, 0)
        !CLMTraditionalMEd = IIf(IsNumeric(A9), A9, 0)
        !CLMphysioTheraphy = IIf(IsNumeric(A10), A10, 0)
        !CLMAccident = IIf(IsNumeric(A11), A11, 0)
        !CLMPRVCheckup = IIf(IsNumeric(A12), A12, 0)
        !CLMVision = IIf(IsNumeric(A13), A13, 0)
        !CLMID = IIf(IsNumeric(A14), A14, 0)
        !CLMTSSTO = IIf(IsNumeric(A15), A15, 0)
        !CLMInjection = IIf(IsNumeric(A16), A16, 0)
        !CLMOTHERS1 = IIf(IsNumeric(A17), A17, 0)
 '       !CLMOTHERS2 = IIf(IsNumeric(A17), A17, 0)
        !CLMOthers1Desc = txtOtherDesc1
 '       !CLMOthers2Desc = txtOtherDesc2
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
        !SPFee = IIf(Len(B1), B1, 0)
        !Medicine = IIf(Len(B2), B2, 0)
        !Dental = IIf(Len(B3), B3, 0)
        !Physio = IIf(Len(B4), B4, 0)
        !Maternity = IIf(Len(B5), B5, 0)
        !PRVCheckup = IIf(Len(B6), B6, 0)
        !Vision = IIf(Len(B7), B7, 0)
    End With
    ClaimNotCover.Update
    
    ClaimNotCover.Close
    Set ClaimNotCover = Nothing
End Sub

Private Sub SaveCase()
Dim CISCAS  As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CISCAS where CASNumber = '" & Trim(CASENumber) & "'"
    CISCAS.Open CISSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    NewCaseNo = CASENumber
    
    With CISCAS
        !ChkItemCat = strBillChkItem
        !Remarks = TxtRemarks
        !ClinicCode = Trim(Left(Cboclinic, IIf(InStr(1, Cboclinic, "-") - 1 < 0, Len(Cboclinic), InStr(1, Cboclinic, "-") - 1)))
        !DateVisit = txtDateVisit
        If isTime(txtVisitTime) Then
            !TimeVisit = txtVisitTime
        End If
        !CASStatus = Mid(cboStatus, 1, 1)
        !Claimability = Mid(cboClaimability, 1, 1)
        !Pending = Mid(cboPending, 1, 1)
        !CASDoctor = Trim(txtDoctor)
        !NatureCLM = Trim(Left(CboNatureClm, IIf(InStr(1, CboNatureClm, "-") - 1 < 0, Len(CboNatureClm), InStr(1, CboNatureClm, "-") - 1)))
        !CASTreatReason1 = Trim(Left(CboReason1, IIf(InStr(1, CboReason1, "-") - 1 < 0, Len(CboReason1), InStr(1, CboReason1, "-") - 1)))
        !CASTreatReason2 = Trim(Left(CboReason2, IIf(InStr(1, CboReason2, "-") - 1 < 0, Len(CboReason2), InStr(1, CboReason2, "-") - 1)))
        !CASTreatReason3 = Trim(Left(CboReason3, IIf(InStr(1, CboReason3, "-") - 1 < 0, Len(CboReason3), InStr(1, CboReason3, "-") - 1)))
        If Mid(cboClaimability, 1, 1) = "R" Then
            Call SaveRepudiation
        End If
      
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
            !CASREPCode3 = Trim(Mid(ChkStr(CboRepudiate3), 1, 7))
            
            !CASREPIssuedBy = Logon
            !CASREPDate = Date
        End With
        CASREP.Update
       
    CASREP.Close
    Set CASREP = Nothing
End Sub

Private Sub EntriesEnable(status As Boolean)
    FrameInvoice.Enabled = status
    FrameCase.Enabled = status
    If Mid(cboClaimability, 1, 1) = "R" Then
        FrameRepudiate.Enabled = True
    End If
End Sub

Private Sub ClearInvoice()
On Error Resume Next
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.TabIndex >= 18 And ctl.TabIndex <= 63 Then
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
    CboNatureClm.Text = ""
    TxtPayee.Text = ""
    Cboclinic.Text = ""
    CboReason1.Text = ""
    CboReason2.Text = ""
    CboReason3.Text = ""
    CboRepudiate1.Text = ""
    CboRepudiate2.Text = ""
    CboRepudiate3.Text = ""
    strBillChkItem = "BILL ACCUTE"
    strMEdChkItem = "MEDICINE ACCUTE"
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
End Sub

Private Sub txtOtherDesc1_LostFocus()
    txtOtherDesc1 = ChkStr(txtOtherDesc1)
End Sub

Private Sub TxtRemarks_LostFocus()
    TxtRemarks = ChkStr(TxtRemarks)
End Sub

Private Sub CheckItemsAmt()
Dim ChkSql As String
Dim ChKItemRec As New ADODB.Recordset

    ChkSql = "SELECT Catcode, Amount FROM CheckItems WHERE CatName = '" & strBillChkItem & _
            "' OR CatName ='" & strMEdChkItem & "'"
    ChKItemRec.Open ChkSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    Do Until ChKItemRec.EOF
        If ChKItemRec!CatCode = "BILL" Then
            strBillChkAmt = ChKItemRec!Amount
        Else
            strMEdChkAmt = ChKItemRec!Amount
        End If
    ChKItemRec.MoveNext
    Loop
    ChKItemRec.Close
    Set ChKItemRec = Nothing

    If IsNumeric(A3) Then
        If CDbl(A3) > CDbl(strMEdChkAmt) Then
            lblMedicine.Visible = True
        Else
            lblMedicine.Visible = False
        End If
    End If
    Call getTotal
    If CDbl(IntInvAmt) > CDbl(strBillChkAmt) Then
        lblTotAmt.Caption = "Excess " & strBillChkItem
        lblTotAmt.Visible = True
    Else
        lblTotAmt.Visible = False
    End If
    If IsNumeric(strVisitLimit) Then
        If CDbl(IntInvAmt) > CDbl(strVisitLimit) Then
             lblVisitLmt.Visible = True
        Else
            lblVisitLmt.Visible = False
        End If
    End If
    
    If IsNumeric(strBalLimit) Then
        If CDbl(IntInvAmt) > CDbl(strBalLimit) Then
            lblBalLimit.Visible = True
        Else
            lblBalLimit.Visible = False
        End If
    End If
    
    ChkSql = "SELECT Catcode, Amount FROM CheckItems WHERE CatCode = 'CSMD'" & _
            " and CatName='" & strCLNContract & "'"
    ChKItemRec.Open ChkSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    If ChKItemRec.RecordCount Then
        If CDbl(IntInvAmt) > CDbl(ChKItemRec!Amount) Then
            lblCLnContract.Caption = "Excess Amount " & strCLNContract
            lblCLnContract.Visible = True
        Else
            lblCLnContract.Visible = False
        End If
    End If
    ChKItemRec.Close
    Set ChKItemRec = Nothing
    If IsDate(txtInvoiceDate) Then
        If (Format(Date, "yyyymmdd") - Format(txtInvoiceDate, "yyyymmdd")) > 45 Then
            Lbl45days.Visible = True
        Else
            Lbl45days.Visible = False
        End If
    End If
End Sub
Private Sub saveActualOthers()
    Dim ClmActOth As New ADODB.Recordset
    Dim strsql As String
 
    If EditFlag = False Then
        ClmActOth.Open "Select INVINdex , * from ClaimInvoiceActualOthers", MediConnCISWS, adOpenKeyset, adLockOptimistic
        ClmActOth.AddNew
        ClmActOth!InvINdex = intIndex
    Else
        strsql = "Select INVINdex , * from ClaimInvoiceActualOthers where INVIndex =" & intIndex
        ClmActOth.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    End If
    
    With ClmActOth
        !CLMMedication1 = IIf(Len(frmINVMedicine.txtMEdication1), frmINVMedicine.txtMEdication1, "")
        !CLMMedication2 = IIf(Len(frmINVMedicine.txtMEdication2), frmINVMedicine.txtMEdication2, "")
        !CLMMedication3 = IIf(Len(frmINVMedicine.txtMEdication3), frmINVMedicine.txtMEdication3, "")
        !CLMProc1 = Trim(Left(frmINVMedicine.CboProc1, IIf(InStr(1, frmINVMedicine.CboProc1, "-") - 1 < 0, Len(frmINVMedicine.CboProc1), InStr(1, frmINVMedicine.CboProc1, "-") - 1)))
        !CLMProc2 = Trim(Left(frmINVMedicine.CboProc2, IIf(InStr(1, frmINVMedicine.CboProc2, "-") - 1 < 0, Len(frmINVMedicine.CboProc2), InStr(1, frmINVMedicine.CboProc2, "-") - 1)))
        !CLMProc3 = Trim(Left(frmINVMedicine.CboProc3, IIf(InStr(1, frmINVMedicine.CboProc3, "-") - 1 < 0, Len(frmINVMedicine.CboProc3), InStr(1, frmINVMedicine.CboProc3, "-") - 1)))
        !CLMProc4 = Trim(Left(frmINVMedicine.CboProc4, IIf(InStr(1, frmINVMedicine.CboProc4, "-") - 1 < 0, Len(frmINVMedicine.CboProc4), InStr(1, frmINVMedicine.CboProc4, "-") - 1)))
        !CLMProc5 = Trim(Left(frmINVMedicine.CboProc5, IIf(InStr(1, frmINVMedicine.CboProc5, "-") - 1 < 0, Len(frmINVMedicine.CboProc5), InStr(1, frmINVMedicine.CboProc5, "-") - 1)))
        !CLMProc6 = Trim(Left(frmINVMedicine.CboProc6, IIf(InStr(1, frmINVMedicine.CboProc6, "-") - 1 < 0, Len(frmINVMedicine.CboProc6), InStr(1, frmINVMedicine.CboProc6, "-") - 1)))
        !CLMProc7 = Trim(Left(frmINVMedicine.CboProc7, IIf(InStr(1, frmINVMedicine.CboProc7, "-") - 1 < 0, Len(frmINVMedicine.CboProc7), InStr(1, frmINVMedicine.CboProc7, "-") - 1)))
    End With
    ClmActOth.Update
    
    ClmActOth.Close
    Set ClmActOth = Nothing
End Sub


